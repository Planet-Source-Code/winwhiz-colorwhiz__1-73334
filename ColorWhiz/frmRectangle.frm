VERSION 5.00
Begin VB.Form frmRectangle 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRectangle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2475
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   3900
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3165
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   0
      Top             =   -15
      Width           =   6780
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   248
         X2              =   248
         Y1              =   56
         Y2              =   112
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   168
         X2              =   240
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   152
         X2              =   152
         Y1              =   56
         Y2              =   112
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   160
         X2              =   240
         Y1              =   48
         Y2              =   48
      End
   End
End
Attribute VB_Name = "frmRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbDown As Boolean
Private nOldX As Integer
Private nOldY As Integer
Dim Picture_Shown As Boolean
Dim WWW As Double
Dim HHH As Double
Dim j%
Dim Filename$
Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()
    '--- Set up the Snap Form to the size of the
    '--- Whole Screen capture we did when we choose
    '--- to Get the Rectangular area Capture.
    '--- The "-2" offset prevents the screen from shifting
    '--- slightly when switchting to display screen capture image.
    
    With Me
        .Left = -2
        .Top = -2
        .Width = Screen.Width + 2
        .Height = Screen.Height + 2
    End With
    
    With Picture1
        .Left = -2
        .Top = -2
        .Height = Me.Height
        .Width = Me.Width
    End With
End Sub

Public Sub ShowPicture(picBitmap As Variant)
    '--- Load the Screen that was Captured into Picture Box
    Load Me
    DoEvents
    Picture1.Picture = picBitmap
    Me.Show
    mbDown = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmChart.Show
End Sub

Private Sub Picture1_DblClick()
    Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- This where we set the Begainning of the Box
'--- that will be Drawn around the Capture Area

    mbDown = (Button = 1)
    
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

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim FileSelect$
    Dim NewImage_Width As Double
    Dim NewImage_Height As Double
    Dim XUpperLeft As Long
    Dim YUpperLeft As Long
    Dim XLowerRight As Long
    Dim YLowerRight As Long
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
        .Width = Abs(Line1.X2 - Line1.X1)
        .Height = Abs(Line2.Y2 - Line2.Y1)
    
        '--- Paint the Captured part of the screen to
        '--- form3 Picture1
        .PaintPicture Picture1, 0, 0, _
            XLowerRight - XUpperLeft, _
            YLowerRight - YUpperLeft, _
            XUpperLeft, YUpperLeft, _
            XLowerRight - XUpperLeft, _
            YLowerRight - YUpperLeft ', opcode
            
        '--- IMPORTANT: DO NOT REMOVE THIS DoEvents! Windows needs to "catchup"
        '--- before can use the "painted" picture.
        DoEvents
        mbDown = False
    End With
    
    '--- Put picture image back in calling form and show it
    With frmChart
        '--- Load selected rectangle image into picture box:
        With .NewImage
            '--- Just to be safe, clear picture before
            '--- loading new image:
            .Picture = LoadPicture()
            .Cls
            .Picture = picRectangle.Image
        FileSelect$ = App.Path & "\Capture\TempRect.bmp"
        SavePicture picRectangle.Image, FileSelect$
        Filename$ = FileSelect$
        'frmChart.mnuCopy.Enabled = True
        'frmChart.mnuPrint.Enabled = True
        'frmChart.mnuSave.Enabled = True
        Picture_Shown = True
        
    '--- Capture the Screen
        
        frmChart.Sample.Picture = LoadPicture(Filename$)
        frmChart.NewImage.Picture = LoadPicture(Filename$)
        NewImage_Width = frmChart.NewImage.Width
        NewImage_Height = frmChart.NewImage.Height
        frmChart.ScaleBar.Value = 100
        frmChart.ScaleBar.Max = 500
        If NewImage_Width * (frmChart.ScaleBar.Max / 100) > 32000 Then frmChart.ScaleBar.Max = (32000 / NewImage_Width) * 100
        If NewImage_Height * (frmChart.ScaleBar.Max / 100) > 32000 Then frmChart.ScaleBar.Max = (32000 / NewImage_Height) * 100
        frmChart.NewImage.Cls
        frmChart.NewImage.Top = 0
        frmChart.NewImage.Left = 0
        Picture_Update
        
        End With
        
        .Show
        frmChart.Cls
        frmChart.ScaleBar.Value = 99
        frmChart.ScaleMeter.Caption = "100%"
        frmChart.ScaleBar.Enabled = True
        frmChart.Mode = Custom
        frmChart.imgPicker.Visible = True
        frmChart.imgPreset.Visible = False
        frmChart.cmdPicker.Visible = False
        frmChart.cmdPreset.Visible = False
        frmChart.fraQBColors.Visible = False
        frmChart.Label8.Visible = False
        frmChart.Picture2.Visible = True
        frmChart.Picture1.Visible = True
        frmChart.lblInfo.Visible = True
        frmChart.IconScroll.Visible = True
        frmChart.VMove.Visible = True
        frmChart.HMove.Visible = True
        frmChart.ScaleBar.Visible = True
        frmChart.ScaleMeter.Visible = True
        '--- unload frmCaptureRectangle
        frmChart.Top = 2400
        If frmChart.Visible = False Then
        frmChart.Visible = True
        End If
        Unload Me
    End With

End Sub
Private Sub Picture_Update()
Dim NewImage_Width As Double
Dim NewImage_Height As Double
frmChart.ScaleMeter.Caption = Mid$(Str$(frmChart.ScaleBar.Value), 2) + "%"
WWW = NewImage_Width * (frmChart.ScaleBar.Value / 100)
If WWW > 4125 Then frmChart.NewImage.Width = 4125 Else frmChart.NewImage.Width = WWW
HHH = NewImage_Height * (frmChart.ScaleBar.Value / 100)
If HHH > 3490 Then frmChart.NewImage.Height = 3490 Else frmChart.NewImage.Height = HHH
frmChart.VMove.Max = HHH - frmChart.NewImage.Height
frmChart.HMove.Max = WWW - frmChart.NewImage.Width
frmChart.VMove.Value = 0
frmChart.HMove.Value = 0
frmChart.NewImage.PaintPicture frmChart.Sample, 0, 0, WWW, HHH, 0, 0, NewImage_Width, NewImage_Height, vbSrcCopy
End Sub


