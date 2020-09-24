Attribute VB_Name = "Mod_Capture"

    
Public XOld As Long
Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type PointAPI
    X As Long
    Y As Long
End Type
 
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
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

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
'API Function Declerations
Public Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Public Declare Function SetICMMode Lib "gdi32" (ByVal hDC As Long, ByVal n As Long) As Long
Public Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hDC As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function ClipCursor Lib "User32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "User32" (lpRect As RECT, pt As PointAPI) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "User32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetClientRect Lib "User32" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long

' Returns a handle to the Desktop window.  The desktop
' window covers the entire screen and is the area on top
' of which all icons and other windows are painted.
Private Declare Function GetDesktopWindow Lib "User32" () As Long

' Returns a handle to the foreground window (the window
' the user is currently working). The system assigns a
' slightly higher priority to the thread that creates the
' foreground window than it does to other threads.
Private Declare Function GetForegroundWindow Lib "User32" () As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleCreatePictureIndirect _
    Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Public Declare Function CreatehalfTonePalette Lib "gdi32" Alias "CreateHalftonePalette" (ByVal hDC As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
    ByVal hDC As Long, ByVal wStartIndex As Long, _
    ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
    As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function UpdateColors Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "User32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetWindowDC Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function UpdateWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

' Constants
Public Const ICM_ON = 2
Public Const ICM_OFF = 1
Public Const ICM_QUERY = 3
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
'Variables
Dim lpbmINFO As BITMAPINFO



Public Function CreateBitmapPicture(ByVal hBmp As Long, _
        ByVal hPal As Long) As Picture

'
Dim r   As Long
Dim Pic As PicBmp
'
' IPicture requires a reference to "Standard OLE Types"
'
Dim IPic          As IPicture
Dim IID_IDispatch As GUID
'
' Fill in with IDispatch Interface ID
'
With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
End With
'
' Fill Pic with the necessary parts.
'
With Pic
    .Size = Len(Pic)          ' Length of structure
    .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
    .hBmp = hBmp              ' Handle to bitmap
    .hPal = hPal              ' Handle to palette (may be null)
End With
'
' Create the Picture object.
r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
'
' Return the new Picture object.
'
Set CreateBitmapPicture = IPic
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal bClient As Boolean, ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture

Dim hDCMemory       As Long
Dim hBmp            As Long
Dim hBmpPrev        As Long
Dim r               As Long
Dim hDCSrc          As Long
Dim hPal            As Long
Dim hPalPrev        As Long
Dim RasterCapsScrn  As Long
Dim HasPaletteScrn  As Long
Dim PaletteSizeScrn As Long
Dim LogPal          As LOGPALETTE
'
' Get the proper Device Context (DC) depending on the value of bClient.
'
If bClient Then
    hDCSrc = GetDC(hWndSrc)       'Get DC for Client area.
Else
    hDCSrc = GetWindowDC(hWndSrc) 'Get DC for entire window.
End If
'
' Create a memory DC for the copy process.
'
hDCMemory = CreateCompatibleDC(hDCSrc)
'
' Create a bitmap and place it in the memory DC.
'
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)
'
' Get the screen properties.
'
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)   'Raster capabilities
HasPaletteScrn = RasterCapsScrn And RC_PALETTE       'Palette support
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) 'Palette size
'
' If the screen has a palette make a copy and realize it.
'
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    '
    ' Create a copy of the system palette.
    '
    LogPal.palVersion = &H300
    LogPal.palNumEntries = 256
    r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
    hPal = CreatePalette(LogPal)
    '
    ' Select the new palette into the memory DC and realize it.
    '
    hPalPrev = SelectPalette(hDCMemory, hPal, 0)
    r = RealizePalette(hDCMemory)
End If
'
' Copy the on-screen image into the memory DC.
'
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
    LeftSrc, TopSrc, vbSrcCopy)
'
' Remove the new copy of the on-screen image.
'
hBmp = SelectObject(hDCMemory, hBmpPrev)
'
' If the screen has a palette get back the
' palette that was selected in previously.
'
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
End If
'
' Release the DC resources back to the system.
'
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
'
' Create a picture object from the bitmap
' and palette handles.
'
Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CaptureScreen() As Picture
Dim hWndScreen As Long
'
' Get a handle to the desktop window.
hWndScreen = GetDesktopWindow()
'
' Capture the entire desktop.
'
With Screen
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            .Width \ .TwipsPerPixelX, .Height \ .TwipsPerPixelY)
End With
End Function

Public Function CaptureForm(frm As Form) As Picture
'
' Capture the entire form.
'
With frm
    Set CaptureForm = CaptureWindow(.hwnd, False, 0, 0, _
            .ScaleX(.Width, vbTwips, vbPixels), _
            .ScaleY(.Height, vbTwips, vbPixels))
End With
End Function

Public Function CaptureClient(frm As Form) As Picture
'
' Capture the client area of the form.
'
With frm
    Set CaptureClient = CaptureWindow(.hwnd, True, 0, 0, _
            .ScaleX(.ScaleWidth, .ScaleMode, vbPixels), _
            .ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
End With
End Function

Public Function CaptureActiveWindow() As Picture
Dim hWndActive As Long
Dim RectActive As RECT
Dim blReturn As Long
'
' Get a handle to the active/foreground window.
' Get the dimensions of the window.
'
hWndActive = GetForegroundWindow()
DoEvents

blReturn = GetWindowRect(hWndActive, RectActive)


'
' Capture the active window.
'
With RectActive
    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
            .Right - .Left, .Bottom - .Top)
End With
End Function

Public Sub PrintPictureToFitPage(Prn As Printer, Pic As Picture)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintPictureToFitPage
'    - Prints a Picture object as big as possible.
'
' Prn
'    - Destination Printer object
'
' Pic
'    - Source Picture object
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim PicRatio     As Double
Dim PrnWidth     As Double
Dim PrnHeight    As Double
Dim PrnRatio     As Double
Dim PrnPicWidth  As Double
Dim PrnPicHeight As Double
Const vbHiMetric As Integer = 8
'
' Determine if picture should be printed in landscape
' or portrait and set the orientation.
'
If Pic.Height >= Pic.Width Then
    Prn.Orientation = vbPRORPortrait   'Taller than wide
Else
    Prn.Orientation = vbPRORLandscape  'Wider than tall
End If
'
' Calculate device independent Width to Height ratio for picture.
'
PicRatio = Pic.Width / Pic.Height
'
' Calculate the dimentions of the printable area in HiMetric.
'
With Prn
    PrnWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbHiMetric)
    PrnHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbHiMetric)
End With
'
' Calculate device independent Width to Height ratio for printer.
'
PrnRatio = PrnWidth / PrnHeight
'
' Scale the output to the printable area.
'
If PicRatio >= PrnRatio Then
    '
    ' Scale picture to fit full width of printable area.
    '
    PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
    PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
Else
    '
    ' Scale picture to fit full height of printable area.
    '
    PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
    PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
End If
'
' Print the picture using the PaintPicture method.
'
Call Prn.PaintPicture(Pic, 0, 0, PrnPicWidth, PrnPicHeight)
End Sub

Public Function GetDesktop(frm As Form)
    Dim HW As Long
    Dim HA As Long
    Dim iLeft As Integer
    Dim iTop As Integer
    Dim iWidth As Integer
    Dim iHeight As Integer
    frm.AutoRedraw = True
    frm.Show
    frm.Hide
    
    DoEvents
    HA = GetDC(GetDesktopWindow())
    iLeft = frm.Left / Screen.TwipsPerPixelX
    iTop = frm.Top / Screen.TwipsPerPixelY
    iWidth = frm.ScaleWidth
    iHeight = frm.ScaleHeight
    Call BitBlt(frm.hDC, 0, 0, iWidth, iHeight, HA, iLeft, iTop, vbSrcCopy)
    frm.Picture = frm.Image
    
    frm.Show

End Function

Public Function Gerade(Number) As Boolean 'Function to see if number is dividable by 2
If Round(Number / 2, 0) = Number / 2 Then
    Gerade = True
Else
    Gerade = False
End If
End Function

Public Sub Pause(Delay)
Dim StartTime
    StartTime = GetTickCount
    Do
    Loop Until StartTime + Delay < GetTickCount
End Sub


