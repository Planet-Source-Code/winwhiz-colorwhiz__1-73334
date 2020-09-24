VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScreen 
   AutoRedraw      =   -1  'True
   Caption         =   "Advanced Print Screen Utility"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7980
   Icon            =   "frmScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRectangle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5955
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   7605
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      Left            =   7650
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":0442
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":089E
            Key             =   ""
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":0CF2
            Key             =   ""
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":1146
            Key             =   ""
            Object.Tag             =   "Capture Rectangle"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":15D0
            Key             =   ""
            Object.Tag             =   "Capture Full Screen"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":1A24
            Key             =   ""
            Object.Tag             =   "About"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":1E78
            Key             =   ""
            Object.Tag             =   "Print image"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":238A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":271C
            Key             =   ""
            Object.Tag             =   "CopyToClipboard"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":2AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScreen.frx":2EF8
            Key             =   ""
            Object.Tag             =   "Crop"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H00808080&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   1
      Top             =   450
      Width           =   7635
      Begin MSComDlg.CommonDialog CD1 
         Left            =   4950
         Top             =   210
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   0
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   151
         TabIndex        =   4
         Top             =   0
         Width           =   2265
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   37
            X2              =   117
            Y1              =   30
            Y2              =   30
         End
         Begin VB.Line Line2 
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   29
            X2              =   29
            Y1              =   38
            Y2              =   94
         End
         Begin VB.Line Line3 
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   45
            X2              =   117
            Y1              =   102
            Y2              =   102
         End
         Begin VB.Line Line4 
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   125
            X2              =   125
            Y1              =   38
            Y2              =   94
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open a new Image."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveIt"
            Object.ToolTipText     =   "Save Image As"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Capture1"
            Object.ToolTipText     =   "Capture Rectangular Area"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Capture2"
            Object.ToolTipText     =   "Capture Full Screen"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Image"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy Image to Clipboard"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crop"
            Object.ToolTipText     =   "Select area to crop"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About Screen Ripper"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnBB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuAA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCrop 
         Caption         =   "Select Crop &Area"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Image to Clipboard"
      End
   End
   Begin VB.Menu mnuB 
      Caption         =   "&Capture"
      Begin VB.Menu mnuRectangle 
         Caption         =   "Capture &Rectangular Area"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Capture Full &Screen"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Capture Color"
      End
   End
   Begin VB.Menu mnuC 
      Caption         =   "&About"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--- Original concept and design by Bob "Real Redneck" Davis (adavis354@comcast.net).
'--- Enhancements by Gary Choma (gchoma@hotmail.com).
'--- Special thanks to www.planet-source-code.com!
'--- If you can make this program better, your welcome to do so...
'--- ...and please then share it!

Private mbCrop As Boolean
Private mbDown As Boolean
Private nOldX As Integer
Private nOldY As Integer

'--- nTitleBarHeight may vary between versions/settings of Windows.
'--- for example, in WinXP, the titlebar is much thicker than in Win95.
Private Const mcnTITLE_BAR_HEIGHT As Integer = 400

'--- Printscreen API declaration:
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetDesktopWindow Lib "User32" () As Long

Private Function CaptureDesktop() As Boolean
'--- 07/30/2004 GCC - Better way to capture screen...(doesn't use Windows clipboard!):
    Dim hWndScreen As Long
    Dim nXpos As Long
    Dim I As Integer
    
    '--- Hide the frmChart so that it will not
    '--- be included in the Screen Capture
    '--- NOTE: It seems sometimes that when Windows gets bogged down,
    '--- it "captures" a ghost of frmScreen because it didn't fully
    '--- hide it.  Here are some attempts to avoid that:
    '--- To make sure this form is not included in screencapture, let's
    '--- hide it AND move it off to the left of the screen, then move it
    '--- back and show it.
    '--- 07/30/2004 GCC - Added loop with DoEvents...seems to have helped.
    nXpos = Me.Left
    Me.Move (Me.Left + Screen.Width), Me.Height
    DoEvents
    '--- Need to issue a .Hide so that focus goes to what's underneath this form.
    Me.Hide
    For I = 1 To 50
        DoEvents
    Next
    
    ' Get a handle to the desktop window.
    hWndScreen = GetDesktopWindow()
    
    ' Capture the entire desktop.
    With Screen
        Set Picture2.Picture = CaptureWindow(hWndScreen, False, 0, 0, _
                Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
    End With
    
    Me.Left = nXpos

    CaptureDesktop = True

End Function

Private Function CaptureDesktopOLD() As Boolean
'--- This where the Screen gets Captured
'--- Captures screenshot image and puts it in Picture2.Picture
    Dim mvContents As Variant
    Dim mnClpFmt As Integer
    Dim I As Long
    Dim nErrorCount As Long
    Dim nXpos As Long
    
    On Error Resume Next
    
    '--- Hide the frmChart so that it will not
    '--- be included in the Screen Capture
    '--- NOTE: It seems sometimes that when Windows gets bogged down,
    '--- it "captures" a ghost of frmScreen because it didn't fully
    '--- hide it.  Here are some attempts to avoid that:
    '--- To make sure this form is not included in screencapture, let's
    '--- hide it AND move it off to the left of the screen, then move it
    '--- back and show it.
    '--- 07/30/2004 GCC - Added loop with DoEvents...seems to have helped.
    nXpos = Me.Left
    Me.Move (Me.Left + Screen.Width), Me.Height
    DoEvents
    '--- Need to issue a .Hide so that focus goes to what's underneath this form.
    Me.Hide
    For I = 1 To 50
        DoEvents
    Next

    '--- Initialize variables
    mnClpFmt = 0
    Set mvContents = Nothing
    With Clipboard
        
        '--- from the VB Help file:
        If .GetFormat(vbCFText) Then mnClpFmt = mnClpFmt + 1
        If .GetFormat(vbCFBitmap) Then mnClpFmt = mnClpFmt + 2
        If .GetFormat(vbCFDIB) Then mnClpFmt = mnClpFmt + 4
        If .GetFormat(vbCFRTF) Then mnClpFmt = mnClpFmt + 8
        
        '--- Cache current contents of clipboard:
        Select Case mnClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                mvContents = .GetText(vbCFText)
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                Set mvContents = .GetData
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                mvContents = .GetData(mnClpFmt)
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                mvContents = .GetText(vbCFRTF)
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
        DoEvents

        On Error GoTo ErrorHandler
    
        '--- Activate Printscreen, which puts screen capture in Clipboard
        Call keybd_event(vbKeySnapshot, 1, 0, 0)
        '--- IMPORTANT: DoEvents are needed to give Windows a chance to
        '--- "keep up / finish up".  It appears that whenever interacting
        '--- programmatically with the Windows Clipboard, judicious use
        '--- of DoEvents are needed surrounding those calls to allow Windows
        '--- to finish processing the relatively time-intensive Clipboar work.
        '--- Otherwise, the program doesn't work...no screen captures show up
        '--- in the Picturebox controls!
        DoEvents
        Picture2.Cls '--- Actually, this seems to help with the processing timing
                     '--- which the DoEvents doesn't seem to be always effective enough?
        Picture2.Picture = .GetData()
        
        
    End With
    
    DoEvents
    CaptureDesktopOLD = True
    
    '--- created from VB help file example.
    On Error Resume Next
    If Not IsEmpty(mvContents) Then
        '--- Restore cached contents of the Windows clipboard
        Select Case mnClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                Clipboard.Clear
                DoEvents
                Clipboard.SetText mvContents, vbCFText
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                Clipboard.Clear
                DoEvents
                Clipboard.SetData mvContents
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                '--- Not sure if this is correct because I'm not sure how
                '--- to set both text and a bitmap into the clipboard
                Clipboard.Clear
                DoEvents
                Clipboard.SetData mvContents
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                '--- i.e. Copied text within MSWord
                Clipboard.Clear
                DoEvents
                Clipboard.SetText mvContents, vbCFRTF
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
    End If
    
    Me.Left = nXpos
    
    Exit Function
ErrorHandler:
    If Err.Number = 521 Then
        Err.Clear
        If nErrorCount < 5 Then
            nErrorCount = nErrorCount + 1
            Resume
        Else
            If MsgBox("Couldn't open Windows Clipboard.  Try again?", vbExclamation + vbYesNo) = vbYes Then
                Resume
            End If
        End If
    Else
        MsgBox "Error number: " & Err.Number & ". " & Err.Description
    End If
    CaptureDesktopOLD = False
    Me.Left = nXpos
    Me.Show
End Function

Private Sub Form_Activate()
    AdjustScrollbars
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub ResizePictureForScrollbars()
    With Picture1
        '--- 07/30/2004 GCC - Size picture differently if scrollbars are hidden.
        If VScroll1.Visible Then
            .Width = Me.Width - VScroll1.Width - 150 - Picture1.Left
        Else
            .Width = Me.Width - 120 - Picture1.Left
        End If
        
        If HScroll1.Visible Then
            .Height = Me.Height - Picture1.Top - Toolbar1.Height - HScroll1.Height - mcnTITLE_BAR_HEIGHT - 30
        Else
            .Height = Me.Height - Picture1.Top - HScroll1.Height - mcnTITLE_BAR_HEIGHT - 30
        End If
    End With
    
    With VScroll1
        '--- 07/30/2004 GCC - Better way to position scrollbar:
        '.Left = Me.Width - VScroll1.Width - 120
        .Left = Picture1.Left + Picture1.Width '+ 15
        .Height = Picture1.Height
        .Top = Picture1.Top
        .Value = 0
    End With
    
    With HScroll1
        '--- 07/30/2004 GCC - Better way to position scrollbar:
        '.Top = Me.Height - HScroll1.Height - mcnTITLE_BAR_HEIGHT - Toolbar1.Height
        .Top = Picture1.Top + Picture1.Height '+ 15
        .Width = Picture1.Width
        .Value = 0
        .Left = Picture1.Left
    End With
    
    '--- Picture2 is contained within Picture1
    With Picture2
        .Left = 0
        .Top = 0
    End With
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    '--- Limit how small user can resize the form:
    With Me
        If .Width < 3000 Then
            .Width = 4000
        End If
        If .Height < 2000 Then
            .Height = 2000
        End If
    End With
    
    ResizePictureForScrollbars
        
    AdjustScrollbars
End Sub

Private Sub mnuAbout_Click()
'--- The about Screen Capture Utility
    Dim ms As String
    ms = "Make sure screen to capture from is directly below this form" & vbCrLf
    MsgBox ms, , Me.Caption
End Sub

Private Sub mnuColor_Click()
frmChart.chkCapture.Value = 1
End Sub

Private Sub mnuCrop_Click()
    Crop
End Sub

Private Sub mnuRectangle_Click()
'--- This is where we start the Capture of
'--- a choosen Rectangular Area

    '--- We capture the Whole Screen even if
    '--- we only want a part of it:
    If CaptureDesktop Then
        '--- Show the Form where all the Work will take place
        DoEvents
        frmRectangle.ShowPicture Picture2.Picture
    End If
End Sub

Private Sub mnuCopy_Click()
    If Picture2.Picture <> 0 Then
        Clipboard.Clear
        DoEvents
        Clipboard.SetData Picture2.Picture
        DoEvents
        MsgBox "Image saved to Windows clipboard.  Use Paste or CTL+V in another application, such as Word, to paste image from clipboard.", vbInformation, "Copy Image to Clipboard"
    Else
        MsgBox "Please capture or load an image before copying to clipboard.", vbInformation, "Nothing To Copy"
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
'--- This is where we choose a Image that is
'--- already on the Disk so that is can be save
'--- by the User
    
    '--- Set the Filters
    With CD1
        .Filter = "GIF Files (*.gif)|*.gif|JPEG Files" & _
                 "(*.jpg)|*.jpg|Bitmap Files (*.bmp)|*.bmp"
        '--- Specify default filter
        .FilterIndex = 2
        '--- set starting Path
        .InitDir = "c:\aaaaaa" 'Path1
        
        .Flags = cdlOFNExplorer
        
        '--- Show the Open Dialog
        .ShowOpen
        '--- If Canceled is Pressed
        If .FileName = "" Then Exit Sub
        '--- Load the Choosen Image to the Picture Box
        Picture2.Picture = LoadPicture(.FileName)
    End With
End Sub

Private Sub mnuPrint_Click()
    If Picture2.Picture <> 0 Then
        frmPrintScreen.PrintBitmap Picture2.Picture
    Else
        MsgBox "Please capture or load an image before printing.", vbInformation, "Nothing To Print"
    End If
End Sub

Private Sub mnuSave_Click()
'--- This where we save the Captured part of the Screen
'--- to disk. I would have set it save as JPG also
'--- but PSC want let You upload the DLL needed
'--- for now save only as a BMP
    If Picture2.Picture = 0 Then
        MsgBox "Please capture or load an image before saving.", vbInformation, "Nothing To Save"
    Else
        '--- Set the Filters
        With CD1
            .Filter = "Bitmap Files (*.bmp)|*.bmp"
            '--- Specify default filter
            .FilterIndex = 2
            '--- Hide the "Open as read only" checkbox when saving.
            .Flags = cdlOFNHideReadOnly
            '--- Show the Open Dialog
            .ShowSave
            '--- If Canceled is Pressed
            If .FileName = "" Then Exit Sub
            '--- Save the Image
            SavePicture Picture2.Image, .FileName
        End With
    End If
End Sub

Private Sub mnuFullScreen_Click()
'--- This is where we capture the FULL screen
'--- when the User chooses to capture the FULL Screen

    '--- Capture the Screen
    CaptureDesktop
        
    Me.Show
    AdjustScrollbars
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'--- Where we do the action for the Clicked Icons
    On Error Resume Next
    Select Case Button.Key
        Case "Exit"
            mnuExit_Click
        Case "Open"
            mnuOpen_Click
        Case "SaveIt"
            mnuSave_Click
        Case "Capture1"
            mnuRectangle_Click
        Case "Capture2"
            mnuFullScreen_Click
        Case "Copy"
            mnuCopy_Click
        Case "Print"
            mnuPrint_Click
        Case "Crop"
            mnuCrop_Click
        Case "About"
            mnuAbout_Click
    End Select
End Sub

Private Sub VScroll1_Change()
'--- Make the changes to the Scrollbars as needed
'--- NOTE: the "+ 60" offset is so that you can see right to
'--- the very edge of the screencapture image.
    '--- 07/30/2004 GCC - Handle very top/ very bottom conditions better:
    Select Case VScroll1.Value
        Case 0
            '--- scroll to very top
            Picture2.Top = 0
            
        Case VScroll1.Max
            '--- scroll to very bottom
            Picture2.Top = 0 - (Picture2.Height - (Picture1.Height / Screen.TwipsPerPixelY) + 4) '((VScroll1.Value / Screen.TwipsPerPixelX) + 4)
        Case Else
            If (Picture2.Height * Screen.TwipsPerPixelY) > Picture1.Height Then
                Picture2.Top = 0 - ((VScroll1.Value / Screen.TwipsPerPixelX) + 4)
            End If
    End Select
End Sub

Private Sub HScroll1_Change()
'--- Make the Changes to the Scrollbars as needed
'--- NOTE: the "+ 60" offset is so that you can see right to
'--- the very edge of the screencapture image.
    '--- 07/30/2004 GCC - Handle very left/ very right conditions better:
    Select Case HScroll1.Value
        Case 0
            '--- scroll to very left
            Picture2.Left = 0
            
        Case HScroll1.Max
            '--- scroll to very right
            Picture2.Left = 0 - (Picture2.Width - (Picture1.Width / Screen.TwipsPerPixelX) + 5)
        Case Else
            If (Picture2.Width * Screen.TwipsPerPixelX) > Picture1.Width Then
                Picture2.Left = 0 - ((HScroll1.Value / Screen.TwipsPerPixelX) + 4)
            End If
    End Select
    
    
End Sub

Private Sub AdjustScrollbars()
'--- Adjust scrollbars' proportions according to size of Picture2
    Dim oActiveControl As Control
    '--- 07/30/2004 GCC - If visible property changes on scroll bars, picture needs resizing.
    Dim bScrollBarVisibleChanged As Boolean
    
    'Exit Sub
    On Error Resume Next
    '--- Cache active control to fix scrollbar bug of blinking scrollbar button having focus
    '--- then not resizing after scrollbar resizes:
    '--- NOTE: Picture2.ScaleMode = pixels, but frmScreen.ScaleMode = TWIPS, so becareful!!
    With VScroll1
        .Value = 0
        .Min = 0
        .Max = IIf(Picture1.Height > (Picture2.ScaleHeight * Screen.TwipsPerPixelX), 0, (Picture2.Height * Screen.TwipsPerPixelY) - Picture1.Height)
        .SmallChange = (.Max / 20) + 1
        .LargeChange = (.Max / 5) + 1
        If .Visible <> (.Max > 0) Then
            bScrollBarVisibleChanged = True
            .Visible = (.Max > 0)
        End If
        .Refresh
    End With
    
    '--- SetUp the HScroll1 Scroolbar
    '--- incase the Image is Larger than the PictureBox
    With HScroll1
        .Value = 0
        .Min = 0
        .Max = IIf(Picture1.Width > (Picture2.Width * Screen.TwipsPerPixelX), 0, (Picture2.Width * Screen.TwipsPerPixelX) - Picture1.Width)
        .SmallChange = (.Max / 20) + 1
        .LargeChange = (.Max / 5) + 1
        If .Visible <> (.Max > 0) Then
            bScrollBarVisibleChanged = True
            .Visible = (.Max > 0)
        End If
        .Refresh
    End With
    
    '--- Some tweaks to get rid of the annoying "flashing" scroll buttons:
    If Picture1.Visible Then
        Set oActiveControl = Me.ActiveControl
        Picture1.SetFocus
        Select Case oActiveControl.Name
            Case "HScroll1", "VScroll1"
                If oActiveControl.Max <> 0 Then
                    'oActiveControl.SetFocus
                End If
            Case Else
                oActiveControl.SetFocus
        End Select
        Set oActiveControl = Nothing
    End If
    
    If bScrollBarVisibleChanged Then
        ResizePictureForScrollbars
    End If

End Sub

Public Sub ActivateRectangle()
    '--- (not used in stand alone version of app)
    '--- Exposed method so that if you include the project forms
    '--- within another application, you can activate the "rectangle picker"
    '--- from another part of the project without showing frmScreen first.
    mnuRectangle_Click
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- This where we set the Begainning of the Box
'--- that will be Drawn around the Capture Area

    If mbCrop Then
        mbDown = (Button = 1)
        Picture2.MousePointer = vbCrosshair
        
        With Line1
            .X1 = X
            .X2 = X
            .Y1 = Y
            .Y2 = Y
        End With
            
        With Line2
            .X1 = X
            .X2 = X
            .Y1 = Y
            .Y2 = Y
        End With
            
        With Line3
            .X1 = X
            .X2 = X
            .Y1 = Y
            .Y2 = Y
        End With
            
        With Line4
            .X1 = X
            .X2 = X
            .Y1 = Y
            .Y2 = Y
        End With
            
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
        
        nOldX = X
        nOldY = Y
    End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- Where we Draw the Box around the Choosen Area as you hold down the Left Mouse
'--- button and Drag in any direction to create a rectangle
    If mbDown Then
        With Line1
            .X1 = nOldX
            .X2 = X
            .Y1 = nOldY
            .Y2 = nOldY
        End With
        
        With Line2
            .X1 = nOldX
            .X2 = nOldX
            .Y1 = nOldY
            .Y2 = Y
        End With
        
        With Line3
            .X1 = X
            .X2 = X
            .Y1 = nOldY
            .Y2 = Y
        End With
        
        With Line4
            .X1 = nOldX
            .X2 = X
            .Y1 = Y
            .Y2 = Y
        End With
    End If

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    Dim XUpperLeft As Long
    Dim YUpperLeft As Long
    Dim XLowerRight As Long
    Dim YLowerRight As Long
        
    '--- 07/30/2004 GCC - Only process this MouseUp event during an actual "crop":
    If mbCrop Then
    
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
        Picture2.MousePointer = vbDefault
        Toolbar1.Buttons(11).Value = tbrUnpressed
        mbCrop = False
        
        '--- Determine the upper left hand corner & lower right hand corner
        '--- XY coordinates.  By doing this, it doesn't matter which
        '--- direction the user "dragged" the rectangle:
        XUpperLeft = Line1.X1
        If Line1.X2 < XUpperLeft Then
            XUpperLeft = Line1.X2
        End If
        With Line2
            If .X1 < XUpperLeft Then
                XUpperLeft = .X1
            End If
            If .X2 < XUpperLeft Then
                XUpperLeft = .X2
            End If
        End With
        
        YUpperLeft = Line1.Y1
        If Line1.Y2 < YUpperLeft Then
            YUpperLeft = Line1.Y2
        End If
        With Line2
            If .Y1 < YUpperLeft Then
                YUpperLeft = .Y1
            End If
            If .Y2 < YUpperLeft Then
                YUpperLeft = .Y2
            End If
        End With
        
        XLowerRight = Line1.X1
        If Line1.X2 > XLowerRight Then
            XLowerRight = Line1.X2
        End If
        With Line2
            If .X1 > XLowerRight Then
                XLowerRight = .X1
            End If
            If .X2 > XLowerRight Then
                XLowerRight = .X2
            End If
        End With
        
        YLowerRight = Line1.Y1
        If Line1.Y2 > YLowerRight Then
            YLowerRight = Line1.Y2
        End If
        With Line2
            If .Y1 > YLowerRight Then
                YLowerRight = .Y1
            End If
            If .Y2 > YLowerRight Then
                YLowerRight = .Y2
            End If
        End With
        
        '--- Selected a single pixel (clicked, no drag)
        If XUpperLeft = XLowerRight Then XLowerRight = XLowerRight + 1
        If YUpperLeft = YLowerRight Then YLowerRight = YLowerRight + 1
    
        '--- Set the picRectangle to the size
        '--- we will paint the Image to
        With picRectangle
            .Picture = LoadPicture()
            .Cls
            DoEvents
            .Width = Abs(Line1.X2 - Line1.X1) * Screen.TwipsPerPixelX
            .Height = Abs(Line2.Y2 - Line2.Y1) * Screen.TwipsPerPixelY
        
            '--- Paint the Captured part of the screen to
            '--- form3 Picture2
            .PaintPicture Picture2.Picture, 0, 0, _
                (XLowerRight - XUpperLeft), _
                (YLowerRight - YUpperLeft), _
                XUpperLeft, YUpperLeft, _
                (XLowerRight - XUpperLeft), _
                (YLowerRight - YUpperLeft)  ', opcode
            
            '--- IMPORTANT: DO NOT REMOVE THIS DoEvents! Windows needs to "catchup"
            '--- before can use the "painted" picture.
            DoEvents
            mbDown = False
        End With
        
        '--- Load selected rectangle image into picture box:
        With Picture2
            '--- Incase picture was scrolled over (via scrollbars), reset it's position
            .Left = 0
            .Top = 0
            '--- Just to be safe, clear picture before
            '--- loading new image:
            .Picture = LoadPicture()
            .Cls
            .Picture = picRectangle.Image
        End With
        AdjustScrollbars
    End If
End Sub

Private Sub Crop()
    mbCrop = True
    Picture2.MousePointer = vbCrosshair
    Toolbar1.Buttons(11).Style = tbrCheck
    Toolbar1.Buttons(11).Value = tbrPressed
End Sub
