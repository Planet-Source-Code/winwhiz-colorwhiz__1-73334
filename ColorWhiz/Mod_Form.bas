Attribute VB_Name = "Mod_Form"

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const LB_ITEMFROMPOINT = &H1A9
Public Enum SoundOps
       SND_SYNC = &H0
      SND_ASYNC = &H1
  SND_NODEFAULT = &H2
       SND_LOOP = &H8
     SND_NOSTOP = &H10
      SND_PURGE = &H40
     SND_NOWAIT = &H2000
     SND_MEMORY = &H4
End Enum
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Sub sndPlay(strName As String, sndType As Long)

 ' plays a sound. :O
 sndPlaySound App.Path & "/Sounds/" & strName & ".wav", sndType

End Sub

Public Function SetWinPos(iPos As Integer, lHWnd As Long) As Boolean
Dim lwinpos As Long
iPos = 1

Select Case iPos
    Case 1
        lwinpos = HWND_TOPMOST
    End Select
If SetWindowPos(lHWnd, lwinpos, 0, 0, 0, 0, SWP_NOMOVE _
                                    + SWP_NOSIZE) Then
SetWinPos = True
End If
End Function
Public Sub SetZeroBorder(frm As Form)
Dim hRgn As Long
Dim fScaleMode As Long
Dim ScrX As Long, ScrY As Long
Dim fLeft As Long, fTop As Long
Dim fBottom As Long, fRight As Long
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    With frm
        fScaleMode = .ScaleMode
        .ScaleMode = 1
        fLeft = (.Width - .ScaleWidth) / 2 / ScrX
        fTop = (.Height - .ScaleHeight) / ScrY - fLeft
        fRight = .Width / ScrX - fLeft
        fBottom = .Height / ScrY - fLeft
        hRgn = CreateRectRgn(fLeft, fTop, fRight, fBottom)
        SetWindowRgn .hwnd, hRgn, True
        .ScaleMode = fScaleMode
        DeleteObject hRgn
    End With
End Sub



      
