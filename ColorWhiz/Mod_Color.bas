Attribute VB_Name = "Mod_Color"
Public gsDatabase As String
Public gConn As ADODB.Connection
'Type Declerations
Public Type RGBTRIPLE
    rgbtBlue As Byte
    rgbtGreen As Byte
    rgbtRed As Byte
End Type

Public Type HSB
    Hue As Single
    Saturation As Single
    Brightness As Single
    End Type
    
Public Type CMYK
    Cyan As Integer
    Magenta As Integer
    Yellow As Integer
    k As Integer
    End Type

Public Type RGB
    Red As Integer
    Green As Integer
    Blue As Integer
    End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAPINFOHEADER       ' 40 bytes
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

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmicolors(15) As Long
End Type

Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry() As PALETTEENTRY
End Type
'API Function Declerations
Public Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Public Declare Function SetICMMode Lib "gdi32" (ByVal hDC As Long, ByVal n As Long) As Long
Public Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hDC As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Function ClientToScreen Lib "User32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ClipCursor Lib "User32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "User32" (lpRect As RECT, pt As PointAPI) As Long

Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "User32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatehalfTonePalette Lib "gdi32" Alias "CreateHalftonePalette" (ByVal hDC As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "User32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Constants
Public Const BitsPixel = 12
Public Const Planes = 14

Public Const BDR_RAISEDINNER = &H4

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Const ICM_ON = 2
Public Const ICM_OFF = 1
Public Const ICM_QUERY = 3

Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

'Variables
Public pMode As Integer
Dim lpbmINFO As BITMAPINFO
Public lpBI As BITMAPINFO
Public m_Color As Long
Public mOldColor As Long
Public SelectBox As RECT
Public MainBox As RECT
Public Preset() As RECT
Public SelectedPos As PointAPI
Public SelectedMainPos As Single
Public cPaletteIndex As Integer
'Public optNotClicked As Boolean
Public svdColor() As Long
Public m_WebColors As Boolean
Public LastColor As Long
Public Sub Main()
    Dim sMsg As String
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    'Check if application is already running
    If App.PrevInstance = True Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "ColorWhiz is already running.", vbOKOnly + vbExclamation, "Order Entry System"
        End
    End If
    
    'Open application database
    Call OpenDatabase
    
    'Show main form
    frmChart.Show
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub OpenDatabase()
    Dim sMsg As String
    Dim sConnect As String
    
    On Error GoTo OPEN_ERROR
    
    gsDatabase = App.Path & "\PantoneConv.mdb"
    sConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & gsDatabase & ";Uid=;Pwd=;"
    Set gConn = New ADODB.Connection
    gConn.ConnectionString = sConnect
    gConn.Open
    
Exit Sub
OPEN_ERROR:
    sndPlay "BOING", SoundOps.SND_ASYNC
    sMsg = "Error opening database '" & gsDatabase & "'.  Error: " & Err.Description
    MsgBox sMsg, vbOKOnly + vbCritical, "System Error"
    End
End Sub




Sub LoadVariantsHue(Red As Integer, Green As Integer, Blue As Integer)
  
    'On Error Resume Next
    Dim St As Long
    
    Dim X As Integer, Y As Integer
    Dim sDc As Long
    Dim K1 As Double, K2 As Double, K3 As Double
    K1 = Red / 255
    K2 = Green / 255
    K3 = Blue / 255
    With frmChart
        .DrawWidth = 1
        .DrawMode = 13
        
        Dim M1    As Double, M2     As Double, M3     As Double
        Dim J1    As Double, J2     As Double, J3     As Double
        Dim YMax As Byte
        Dim shdBitmap(0 To 196608) As Byte  '256 ^ 2 * 3
        Dim L As Long
        Dim bpos As Long
        Dim count As Long
        bpos = 0
        count = 0
        
        With lpBI.bmiHeader
            .biHeight = 256
            .biWidth = 256
        End With
        
        On Error Resume Next
        For Y = 255 To 0 Step -1
                 M1 = Red - Y * K1
                 M2 = Green - Y * K2
                 M3 = Blue - Y * K3
                 YMax = 255 - Y
                 J1 = (YMax - M1) / 255
                 J2 = (YMax - M2) / 255
                 J3 = (YMax - M3) / 255
            For X = 255 To 0 Step -1
                If m_WebColors Then
                    shdBitmap(bpos) = CInt((M3 + X * J3) / &H33) * &H33
                    shdBitmap(bpos + 1) = CInt((M2 + X * J2) / &H33) * &H33
                    shdBitmap(bpos + 2) = CInt((M1 + X * J1) / &H33) * &H33
                Else
                    shdBitmap(bpos) = M3 + X * J3   'Blue
                    shdBitmap(bpos + 1) = M2 + X * J2 'Green
                    shdBitmap(bpos + 2) = M1 + X * J1 'Red
                End If
                bpos = bpos + 3
            Next X
        Next Y
        
    BltBitmap frmChart.hDC, shdBitmap, 138, 10, 256, 256, True
       
        
    End With
    frmChart.DrawSelFrame
End Sub




Sub LoadVariantsBrightness(ByVal Brightness As Single)
Dim OldP As PointAPI
Dim v As Integer
On Error Resume Next
Dim H, m As Single
Dim a As Integer, b As Integer, C As Integer, D As Integer, E As Integer, f As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
H = SelectBox.Bottom - SelectBox.Top
m = H / 6
a = m
b = 2 * m
C = 3 * m
D = 4 * m
E = 5 * m
f = 6 * m
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3

With frmChart
    .DrawMode = 13
    sDc = .hDC
End With

With lpBI.bmiHeader
    .biHeight = 256
    .biWidth = 256
End With

    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficciency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    Dim pos As Long
    pos = 0
    
Dim X  As Integer, Y As Integer
For Y = 255 To 0 Step -1
        MV = 1 - Y / 255 ' ""
    '1
        For X = 0 To a
            v = X * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 0) = CInt(Kc * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = Y * Brightness
                bBitmap(pos + 0) = Kc * Brightness
            End If
            pos = pos + 3
        Next X
    '2
        For X = a + 1 To b
            v = Maa - 6 * X ' 255 - (X - A) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 0) = 255 * Brightness
            Else
                bBitmap(pos + 2) = Kc * Brightness
                bBitmap(pos + 1) = Y * Brightness
                bBitmap(pos + 0) = 255 * Brightness
            End If
            pos = pos + 3
        Next X
     '3
        For X = b + 1 To C
            v = (X - b - 1) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 0) = 255 * Brightness
            Else
                bBitmap(pos + 2) = Y * Brightness
                bBitmap(pos + 1) = Kc * Brightness
                bBitmap(pos + 0) = 255 * Brightness
            End If
            pos = pos + 3
        Next X
     '4
        For X = C + 1 To D
            v = Mcc - 6 * X
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = CInt(Kc * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = Y * Brightness
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = Kc * Brightness
            End If
            pos = pos + 3
        Next X
    '5
        For X = D + 1 To E
            v = (X - D - 1) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = CInt(Y * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc * Brightness
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = Y * Brightness
            End If
            pos = pos + 3
        Next X
    '6
        For X = E + 1 To f
            v = Mee - 6 * X
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 0) = CInt(Y * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = Kc * Brightness
                bBitmap(pos + 0) = Y * Brightness
            End If
            pos = pos + 3
        Next X
       
Next Y

    BltBitmap frmChart.hDC, bBitmap, 393, 10, -256, 256, True
    frmChart.DrawSelFrame
    
End Sub


Sub LoadVariantsSaturation(ByVal Saturation As Single)
Dim OldP As PointAPI
Dim v As Integer
On Error Resume Next
Dim H, m As Single
Dim X As Integer, Y As Integer
Dim a As Integer, b As Integer, C As Integer, D As Integer, E As Integer, f As Integer
Dim sDc As Long
Dim Color As Long
Dim Red As Integer, Green As Integer, Blue As Integer
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3
Dim cpos As Long
cpos = 0
H = SelectBox.Bottom - SelectBox.Top
m = H / 6
a = m
b = 2 * m
C = 3 * m
D = 4 * m
E = 5 * m
f = 6 * m

'frmChart.DrawMode = 6
'frmChart.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle

With frmChart
    .DrawWidth = 1
    .DrawMode = 13
    sDc = .hDC
End With
    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficiency
    Mcc = 255 + 6 * C  ' ""
    Mee = 255 + 6 * E  ' ""
    
For Y = 255 To 0 Step -1
        MV = 1 - Y / 255  ' ""
        YPos = SelectBox.Top + Y
    '1
        For X = 0 To a
            v = X * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 0) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y)
                bBitmap(pos + 1) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 0) = Kc + (255 - Y - Kc) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '2
        For X = a + 1 To b
            v = Maa - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 1) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y)
            End If
            pos = pos + 3
        Next X
     '3
        For X = b + 1 To C
            v = (X - b - 1) * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 1) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 1) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y)
            End If
            pos = pos + 3
        Next X
     '4
        For X = C + 1 To D
            v = Mcc - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 0) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 1) = 255 - Y
                bBitmap(pos + 0) = Kc + (255 - Y - Kc) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '5
        For X = D + 1 To E
            v = (X - D - 1) * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 1) = 255 - Y
                bBitmap(pos + 0) = (255 - Y) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '6
        For X = E + 1 To f
            v = Mee - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 1) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 - Y
                bBitmap(pos + 1) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
       
Next Y
    BltBitmap frmChart.hDC, bBitmap, 393, 10, -256, 256, True
    frmChart.DrawSelFrame
    'frmChart.DrawMode = 6
    'frmChart.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle

End Sub

Sub GetRGB(ByRef cl As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim C As Long
    C = cl
    Red = C Mod &H100
    C = C \ &H100
    Green = C Mod &H100
    C = C \ &H100
    Blue = C Mod &H100
End Sub

Sub DrawSlider(ByVal Position As Integer)

    frmChart.DrawMode = 6
    frmChart.DrawWidth = 2
    frmChart.Line (MainBox.Right + 2, Position)-(MainBox.Right + 5, Position)
    frmChart.Line (MainBox.Left - 2, Position)-(MainBox.Left - 5, Position)
    frmChart.DrawWidth = 1
End Sub

Sub LoadSafePalette()
frmChart.FillStyle = 0
frmChart.DrawMode = 13
frmChart.DrawWidth = 1
On Error Resume Next
Dim I, j, k As Integer
Dim L As Long
Dim count As Integer
Dim Plt As Long
Dim ret As Long
Dim br As Long
Dim pal As Long, oldpal As Long
pal = CreatehalfTonePalette(frmChart.hDC)
oldpal = SelectPalette(frmChart.hDC, pal, 0)
RealizePalette (frmChart.hDC)

For I = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            DrawSafeColor Preset(count), I, j, k
        Next k
    Next j
Next I


For I = 217 To 224
    frmChart.FillColor = 0
    Rectangle frmChart.hDC, Preset(I).Left, Preset(I).Top, Preset(I).Right, Preset(I).Bottom
Next I

SelectPalette frmChart.hDC, oldpal, 0
DeleteObject pal

frmChart.DrawSafePicker cPaletteIndex, False
Dim r As Integer, g As Integer, b As Integer
frmChart.GetSafeColor cPaletteIndex, r, g, b
'frmChart.lblSelColor.BackColor = RGB(r, g, b)
End Sub

Sub LoadCustomColors()
    Dim FileHandle As Integer
    Dim I As Integer
    Dim strColor As String
    On Error Resume Next
    FileHandle = FreeFile()
    ReDim svdColor(0 To 224)
    Open App.Path & "/usercolors.cps" For Input As #FileHandle
    I = 0
    frmChart.Cls
    'frmChart.PrintLastColor
    frmChart.FillStyle = 0
    frmChart.DrawMode = 13
    frmChart.DrawWidth = 1
    For I = 0 To 224
        Line Input #FileHandle, strColor
        svdColor(I) = Val(strColor)
        frmChart.ForeColor = vbBlack 'svdColor(i)
        frmChart.FillColor = svdColor(I)
        Rectangle frmChart.hDC, Preset(I).Left, Preset(I).Top, Preset(I).Right, Preset(I).Bottom
    Next I
    Close #FileHandle
    frmChart.DrawSafePicker cPaletteIndex, False
    frmChart.PSet (-100, -100)
End Sub
Sub LoadLastColor()
    Dim Last As Integer
    Dim I As Integer
    Dim strColor As String
    On Error Resume Next
    Last = FreeFile()
    Open App.Path & "/lastcolor.cps" For Input As #Last
        Line Input #Last, strColor
        frmChart.Label2.BackColor = Val(strColor)
    Close #Last
    
End Sub
Sub SaveCustomColors()
    
    Dim FileHandle As Integer
    Dim I As Integer
    On Error Resume Next
    FileHandle = FreeFile()
    Open App.Path & "/usercolors.cps" For Output As #FileHandle
    For I = 0 To 224
        Print #FileHandle, svdColor(I)
    Next I
    Close #FileHandle
  
End Sub



Public Sub LoadMainSaturation(ByVal hDC As Long, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer, ByVal Brightness As Single)

    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    Dim f As Single
    For Y = 0 To 255
        f = 1 - (Y / 255)
        r = Red * f + Y
        g = Green * f + Y
        b = Blue * f + Y

        For X = 0 To 15
            If m_WebColors Then
                bBitmap(pos) = CInt(b * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(g * Brightness / &H33) * &H33
                bBitmap(pos + 2) = CInt(r * Brightness / &H33) * &H33
            Else
                bBitmap(pos) = b * Brightness
                bBitmap(pos + 1) = g * Brightness
                bBitmap(pos + 2) = r * Brightness
            End If
            pos = pos + 3
        Next X
    Next Y

    BltBitmap hDC, bBitmap, 405, 265, 15, -256, True
    DrawMainFrame
End Sub

Public Sub LoadMainHue()
    Dim OldP As PointAPI
    Dim v As Integer
    On Error Resume Next
    Dim H As Single, m As Single
    Dim a As Single, b As Single, C As Single, D As Single, E As Single, f As Single
    Dim Ratio As Single
    
    H = SelectBox.Bottom - SelectBox.Top
    m = H / 6
    a = m
    b = 2 * m
    C = 3 * m
    D = 4 * m
    E = 5 * m
    f = 6 * m
    Dim sBitmap(0 To 16 * 256 * 3) As Byte            '256 ^ 2 * 3
    Dim cpos  As Long
    With lpBI.bmiHeader
        .biHeight = 256
        .biWidth = 15
    End With

    cpos = 0
        For Y = 0 To Int(a)
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = CInt(Y * 6 / &H33) * &H33
                Else
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = Y * 6
                End If
                cpos = cpos + 3
            Next j
        Next Y
    '2
                
        For Y = Int(a) + 1 To Int(b)
            v = 255 - (Y - a) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = 255
                Else
                    sBitmap(cpos + 2) = v
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = 255
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
     '3
         
        For Y = Int(b) + 1 To Int(C)
            v = (Y - b) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 0) = 255
                Else
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = v
                    sBitmap(cpos + 0) = 255
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
     '4
        For Y = Int(C) + 1 To Int(D)
            v = 255 - (Y - C) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = CInt(v / &H33) * &H33
                Else
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = v
                End If
                cpos = cpos + 3
            Next j
        Next Y
    '5
        For Y = Int(D) + 1 To Int(E)
            v = (Y - D) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = 0
                Else
                    sBitmap(cpos + 2) = v
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = 0
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
    '6
        For Y = Int(E) + 1 To Int(f)
            v = 255 - (Y - E) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 0) = 0
                Else
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = v
                    sBitmap(cpos + 0) = 0
                End If
                cpos = cpos + 3
            Next j
        Next Y
        BltBitmap frmChart.hDC, sBitmap, MainBox.Left, MainBox.Bottom, 15, -256, True
        DrawMainFrame
        'frmChart.DrawPicker
End Sub


Public Sub LoadMainBrightness(ByVal hDC As Long, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    
    For Y = 0 To 255
        r = Red - Red * Y / 255
        g = Green - Green * Y / 255
        b = Blue - Blue * Y / 255
        For X = 0 To 15
            If m_WebColors Then
                bBitmap(pos) = CInt(b / &H33) * &H33
                bBitmap(pos + 1) = CInt(g / &H33) * &H33
                bBitmap(pos + 2) = CInt(r / &H33) * &H33
            Else
                bBitmap(pos) = b
                bBitmap(pos + 1) = g
                bBitmap(pos + 2) = r
            End If
            pos = pos + 3
        Next X
    Next Y
    
    BltBitmap hDC, bBitmap, 405, 265, 15, -256, True
    DrawMainFrame
End Sub
   
Public Function HSBtoRGB(hs As HSB, ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
    Dim I As Integer
    Dim f As Single, p As Single, q As Single, t As Single
    hs.Saturation = hs.Saturation * 255 / 100
    hs.Brightness = hs.Brightness * 255 / 100
    If (hs.Saturation = 0) Then
        ' achromatic (grey)
        r = hs.Brightness
        g = r
        b = r
        Exit Function
    End If
    
    hs.Hue = hs.Hue / 60         ' sector 0 to 5
    I = Int(hs.Hue)
    f = hs.Hue - I         ' factorial part of hs.Hue
    p = hs.Brightness * (1 - hs.Saturation / 255)
    q = hs.Brightness * (1 - (hs.Saturation / 255) * f)
    t = hs.Brightness * (1 - (hs.Saturation / 255) * (1 - f))
    Select Case I
        Case 0
            r = hs.Brightness
            g = t
            b = p
        Case 1
            r = q
            g = hs.Brightness
            b = p
        Case 2
            r = p
            g = hs.Brightness
            b = t
        Case 3
            r = p
            g = q
            b = hs.Brightness
        Case 4
            r = t
            g = p
            b = hs.Brightness
        Case Else        ' case 5:
            r = hs.Brightness
            g = p
            b = q
        End Select
    
End Function

Public Function RGBtoHSB(ByVal Color As Long) As HSB
    Dim LargestValue As Integer
    Dim SmallestValue As Integer
    Dim Red  As Integer, Green As Integer, Blue As Integer
    Dim RedRatio As Single, GreenRatio As Single, BlueRatio As Single
    GetRGB Color, Red, Green, Blue
    LargestValue = IIf(Red >= Green, Red, Green)
    LargestValue = IIf(LargestValue >= Blue, LargestValue, Blue)
    SmallestValue = IIf(Red <= Green, Red, Green)
    SmallestValue = IIf(SmallestValue <= Blue, SmallestValue, Blue)
    RGBtoHSB.Brightness = LargestValue * 100 / 255
    If LargestValue <> 0 Then
        RGBtoHSB.Saturation = 100 - (SmallestValue * 100 / LargestValue)
    Else
        RGBtoHSB.Saturation = 0
    End If
    If RGBtoHSB.Saturation = 0 Then
        RGBtoHSB.Hue = 0
    Else
        RedRatio = (LargestValue - Red) / (LargestValue - SmallestValue)
        GreenRatio = (LargestValue - Green) / (LargestValue - SmallestValue)
        BlueRatio = (LargestValue - Blue) / (LargestValue - SmallestValue)
        Select Case LargestValue
        Case Red
            RGBtoHSB.Hue = BlueRatio - GreenRatio
        Case Green
            RGBtoHSB.Hue = (2 + RedRatio) - BlueRatio
        Case Blue
            RGBtoHSB.Hue = (4 + GreenRatio) - RedRatio
        End Select
        RGBtoHSB.Hue = RGBtoHSB.Hue * 60
        If RGBtoHSB.Hue < 0 Then
            RGBtoHSB.Hue = RGBtoHSB.Hue + 360
        End If
    End If
    

End Function

Sub DrawMainFrame()
    Dim MainFrame As RECT
     
    MainFrame.Left = MainBox.Left - 1
    MainFrame.Top = MainBox.Top - 1
    MainFrame.Right = MainBox.Right + 1
    MainFrame.Bottom = MainBox.Bottom + 3
    DrawEdge frmChart.hDC, MainFrame, BDR_SUNKENOUTER, BF_RECT
    frmChart.PSet (-100, -100)
End Sub


Public Function RGBtoCMYK(ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer) As CMYK
    With RGBtoCMYK
    .k = 100
        .Cyan = ((255 - Red) / 255) * 100
        If .k > .Cyan Then
        .k = .Cyan
        End If
        .Magenta = ((255 - Green) / 255) * 100
        If .k > .Magenta Then
        .k = .Magenta
        End If
        .Yellow = ((255 - Blue) / 255) * 100
        If .k > .Yellow Then
        .k = .Yellow
        End If
        .k = IIf(.Cyan < .Magenta, .Cyan, .Magenta)
        If .Yellow < .k Then
            .k = .Yellow
        End If
        If .k > 0 Then
        .k = .k
        .Cyan = .Cyan - .k
        .Magenta = .Magenta - .k
        .Yellow = .Yellow - .k
End If

    End With
    
        
End Function



Public Sub DrawSafeColor(sPos As RECT, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)

    Dim bBitmap(3 * 16 * 16) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    Dim Width As Integer
    Dim Height As Integer
    Width = sPos.Bottom - sPos.Top
    Height = sPos.Bottom - sPos.Top
    For Y = 0 To Height
        For X = 0 To Width
                bBitmap(pos) = Blue
                bBitmap(pos + 1) = Green
                bBitmap(pos + 2) = Red
                pos = pos + 3
        Next X
    Next Y
    BltBitmap frmChart.hDC, bBitmap, sPos.Left, sPos.Top, Width, Height, False


End Sub

Public Sub BltBitmap(ByVal hDC As Long, bmptr() As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal CreatehfTPalette As Boolean)
        Dim lpBI As BITMAPINFO
        lpBI.bmiHeader.biBitCount = 24
        lpBI.bmiHeader.biCompression = BI_RGB
        lpBI.bmiHeader.biWidth = Abs(Width)
        lpBI.bmiHeader.biHeight = Abs(Height)
        lpBI.bmiHeader.biPlanes = 1
        lpBI.bmiHeader.biSize = 40
        If CreatehfTPalette Then
            Dim pal As Long, oldpal As Long
            pal = CreatehalfTonePalette(hDC)
            oldpal = SelectPalette(hDC, pal, 0)
            RealizePalette (hDC)
        End If
        StretchDIBits hDC, X, Y, Width, Height, 0, 0, Abs(Width), Abs(Height), bmptr(0), lpBI, DIB_RGB_COLORS, vbSrcCopy
        If CreatehfTPalette Then
            SelectPalette hDC, oldpal, 0
            DeleteObject pal
        End If
End Sub

