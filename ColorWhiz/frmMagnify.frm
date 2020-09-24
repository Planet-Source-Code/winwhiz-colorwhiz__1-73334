VERSION 5.00
Begin VB.Form frmMagnify 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ControlBox      =   0   'False
   DrawWidth       =   5
   ForeColor       =   &H00A87856&
   LinkTopic       =   "Form1"
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   86
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   3150
      Top             =   2655
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   240
      Left            =   630
      Top             =   735
      Width           =   240
   End
   Begin VB.Label lblCoord 
      BackColor       =   &H00A87856&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   -30
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "frmMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************************************
'  MouseMagnify constants
'**************************************************************************************************
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const MOUSE_BUFFER = 300
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'**************************************************************************************************
'  MouseMagnify structs
'**************************************************************************************************
Private Type PointAPI
     X As Long
     Y As Long
End Type ' POINTAPI

'**************************************************************************************************
'  MouseMagnify Win32API
'**************************************************************************************************
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, _
     ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
     ByVal dwRop As Long) As Long

'**************************************************************************************************
'  MouseMagnify Timer...freakin' simple
'**************************************************************************************************
Private Sub tmrUpdate_Timer()
     Dim m_Cursor As PointAPI
     Dim m_hDC As Long
     Dim lRtn As Long
     Dim X As Long
     Dim Y As Long
     Dim lScrHt As Long
     Dim lScrWt As Long
     Dim Wsize, Hsize As Long
     Cls
     lScrHt = Screen.Height \ Screen.TwipsPerPixelY
     lScrWt = Screen.Width \ Screen.TwipsPerPixelX
     ' Get the position of the mouse cursor
     GetCursorPos m_Cursor
     ' update coordinates label
     lblCoord = "  X = " & m_Cursor.X & "  Y = " & m_Cursor.Y
     ' convert x and y positions into twips and add buffer.  Buffer necessary
     ' to create some space from the mouse cursor so that we don't see the corner
     ' of the zoom box in the magnification.
     ' If we are at the right of the screen
     If m_Cursor.X + (Me.Width \ Screen.TwipsPerPixelX) + (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt And _
          m_Cursor.Y + (Me.Height \ Screen.TwipsPerPixelY) + (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          X = (m_Cursor.X * Screen.TwipsPerPixelX) - (Me.Width + MOUSE_BUFFER)
          Y = (m_Cursor.Y * Screen.TwipsPerPixelY) - (Me.Height + MOUSE_BUFFER)
     ElseIf m_Cursor.X + (Me.Width \ Screen.TwipsPerPixelX) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt Then
          X = (m_Cursor.X * Screen.TwipsPerPixelX) - (Me.Width + MOUSE_BUFFER)
          Y = m_Cursor.Y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     ElseIf m_Cursor.Y + (Me.Height \ Screen.TwipsPerPixelY) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          X = m_Cursor.X * Screen.TwipsPerPixelX + MOUSE_BUFFER
          Y = (m_Cursor.Y * Screen.TwipsPerPixelY) - (Me.Height + MOUSE_BUFFER)
     Else
          X = m_Cursor.X * Screen.TwipsPerPixelX + MOUSE_BUFFER
          Y = m_Cursor.Y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     End If
     ' move the form with the cursor
     Me.Move X, Y, Me.Width, Me.Height
     ' Get the screen device context
     m_hDC = GetWindowDC(0)
     ' Blit the coordinates, passed in the api call, and stretch it into
     ' our form
     Wsize = ScaleWidth * 8
     Hsize = ScaleHeight * 8
     StretchBlt Me.hDC, 0, 0, Wsize, Hsize, _
        m_hDC, m_Cursor.X - 24 / 8, m_Cursor.Y - 24 / 8, 48, 48, vbSrcCopy
     ' Draw a box to make the form distinguishable from the background.  Set the forms
     ' forecolor to make changes to it.
     frmMagnify.Line (0, 0)-(frmMagnify.ScaleWidth + 1, frmMagnify.ScaleHeight + 1), , B
     ' Bring the window to the top.
     'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
     ' release the screen's device context
     lRtn = ReleaseDC(0, m_hDC)
     ' If at coordinate 0, 0 then quit
     If m_Cursor.X = 0 And m_Cursor.Y = 0 Then
          Unload Me
          Set frmMagnify = Nothing
     End If
End Sub ' tmrUpdate_Timer
