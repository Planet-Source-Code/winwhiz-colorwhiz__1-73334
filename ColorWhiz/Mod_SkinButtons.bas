Attribute VB_Name = "Mod_SkinButtons"
'=============================================================
'=============================================================
'            [ Auther  : 'Jim Jose '           ]
'            [ Email   : jimjosev33@yahoo.com  ]
'            [ Created : 09/03/2005            ]
'=============================================================
'            [ Project : 'Skin Button'         ]
'            [ Page    : 'Not Set              ]
'=============================================================
'             'Please do not modify this Title'
'=============================================================
Option Explicit

'[Enums]
Public Enum ButtonStyleEnum
    Styl_None = 0
    Styl_Normal = 1
    Styl_Dissolve = 2
    Styl_Invert = 3
End Enum

'[APIs]
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

' [ Skin Button ]
'=============================================================
Public Sub SkinButtons(frm As Form, Optional BtnStyle As ButtonStyleEnum = 1, Optional ByVal BorderColor As Long = 0, _
                                    Optional ByVal Curvature As Long = 6, Optional ByVal BorderWidth As Long = 10)
Dim X As Long
Dim fWidth As Long, fHeight As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim PicDraw As PictureBox, Ctl As Control
Dim hRgn As Long, TempRgn As Long, hBrush As Long
Dim Rincr As Double, Gincr As Double, Bincr As Double
If BtnStyle = Styl_None Then Exit Sub

    For Each Ctl In frm
        If TypeName(Ctl) = "CommandButton" Then
            If BtnStyle = Styl_Dissolve Then BorderColor = Ctl.Container.BackColor
            fWidth = Ctl.Width / Screen.TwipsPerPixelX:
            fHeight = Ctl.Height / Screen.TwipsPerPixelY
            If BtnStyle = Styl_Invert Then
                GetRGB Ctl.BackColor, R1, G1, B1: GetRGB BorderColor, R2, G2, B2
            Else
                GetRGB BorderColor, R1, G1, B1: GetRGB Ctl.BackColor, R2, G2, B2
            End If
            Rincr = (R2 - R1) / BorderWidth: Gincr = (G2 - G1) / BorderWidth: Bincr = (B2 - B1) / BorderWidth
            Set PicDraw = frm.Controls.Add("vb.picturebox", "Pic_" & Ctl.hwnd): PicDraw.AutoRedraw = True
            PicDraw.Move 0, 0, Ctl.Width, Ctl.Height: PicDraw.BorderStyle = 0
            For X = 0 To BorderWidth
                hRgn = CreateRoundRectRgn(X, X, fWidth - X, fHeight - X, Curvature, Curvature)
                hBrush = CreateSolidBrush(RGB(R1 + X * Rincr, G1 + X * Gincr, B1 + X * Bincr))
                FillRgn PicDraw.hdc, hRgn, hBrush
                If Not X = BorderWidth Then DeleteObject hRgn: DeleteObject hBrush
            Next X
            TempRgn = CreateRoundRectRgn(0, 0, 2 * fWidth, 2 * fHeight, 0, 0)
            CombineRgn hRgn, TempRgn, hRgn, 3
            SetWindowRgn PicDraw.hwnd, hRgn, True
            DeleteObject TempRgn: DeleteObject hBrush
            TempRgn = CreateRoundRectRgn(0, 0, fWidth, fHeight, Curvature, Curvature)
            SetWindowRgn Ctl.hwnd, TempRgn, True
            SetParent PicDraw.hwnd, Ctl.hwnd
            PicDraw.Visible = True: PicDraw.Enabled = False: Set PicDraw = Nothing
        End If
    Next Ctl
End Sub
'=============================================================

'[ Gets the RGB values ]
'=============================================================
Public Sub GetRGB(ByVal LngCol As Long, r As Long, g As Long, b As Long)
  If LngCol < 0 Then LngCol = GetSysColor(15)
  r = LngCol Mod 256
  g = (LngCol And vbGreen) / 256 'Green
  b = (LngCol And vbBlue) / 65536 'Blue
End Sub
'=============================================================

