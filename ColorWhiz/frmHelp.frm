VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " ColorWhiz Help"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2700
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmHelp.frx":076A
   ScaleHeight     =   5970
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   60
      ScaleHeight     =   3465
      ScaleWidth      =   2565
      TabIndex        =   2
      Top             =   1560
      Width           =   2595
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3405
         Left            =   30
         ScaleHeight     =   3375
         ScaleWidth      =   2475
         TabIndex        =   3
         Top             =   30
         Width           =   2505
         Begin VB.PictureBox LContainer 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3465
            Left            =   0
            ScaleHeight     =   231
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   158
            TabIndex        =   4
            Top             =   360
            Width           =   2365
            Begin VB.PictureBox Info 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   12075
               Left            =   60
               ScaleHeight     =   805
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   155
               TabIndex        =   5
               Top             =   60
               Width           =   2325
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SAMPLE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   60
            Width           =   2115
         End
      End
   End
   Begin VB.Timer timerScroll 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   60
      Top             =   5340
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   983
      TabIndex        =   1
      Top             =   5100
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   353
      TabIndex        =   0
      Text            =   "Topics.........."
      Top             =   1140
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   180
      Picture         =   "frmHelp.frx":1C69
      Top             =   2160
      Visible         =   0   'False
      Width           =   2145
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdStop_Click()
If timerScroll.Enabled = True Then
timerScroll.Enabled = False
cmdStop.Caption = "Start"
Else
cmdStop.Caption = "Stop"
timerScroll.Enabled = True
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Using ColorWhiz" Then
    Image1.Visible = True
    frmHelp1.Visible = False
    frmHelp2.Visible = False
    frmHelp3.Visible = False
    frmHelp4.Visible = False
    frmInfo.Visible = False
    ElseIf Combo1.Text = "   Controls Overview." Then
    frmHelp1.Visible = True
    frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/main1.jpg")
    frmHelp2.Visible = False
    frmHelp3.Visible = False
    frmHelp4.Visible = False
    frmInfo.Visible = False
    Image1.Visible = False
    timerScroll.Enabled = False
    Picture1.Visible = False
    cmdStop.Visible = False
    ElseIf Combo1.Text = "   Using image files" Then
        Image1.Visible = True
        frmHelp1.Visible = False
        frmHelp2.Visible = True
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
        frmHelp1.Label1.Caption = "   Move your mouse pointer over a control."
    ElseIf Combo1.Text = "       Load file." Then
        frmHelp1.Visible = False
        frmHelp2.Visible = True
        frmHelp2.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/filelist1.jpg")
        frmHelp2.Label1.Caption = "   To load a file in the Image Viewer, Use the Drive and Folder list boxes to navigate to a specific folder on chosen drive. When you click a image file name, that image is loaded into the main viewer."
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Open file." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/openfile.jpg")
        frmHelp1.Label1.Caption = "   To open an image file for display in the Image Viewer, click on the file open button."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Save file as." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/saveas.jpg")
        frmHelp1.Label1.Caption = "   Allows you to save a loaded image file as a .bmp,.jpeg,.gif,.tiff or .png file."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        cmdStop.Visible = True
        Label2.Caption = "Common Formats"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipFormat.jpg")
    ElseIf Combo1.Text = "       Copy file." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/copy.jpg")
        frmHelp1.Label1.Caption = "   Copies image file to clipboard."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Print image file." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/printing.jpg")
        frmHelp1.Label1.Caption = "   Prints loaded image file."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Thumbnail Viewer." Then
        frmHelp1.Visible = False
        frmHelp2.Visible = True
        frmHelp2.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/thumbnailview.jpg")
        frmHelp2.Label1.Caption = "   Thumbnail Viewer, gives a preview of image files in the selected folder on your computer, clicking on an image in the Thumbnail Viewer loads that image into the main viewer."
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Zoom Contol." Then
        frmHelp1.Visible = False
        frmHelp2.Visible = True
        frmHelp2.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/Scale.jpg")
        frmHelp2.Label1.Caption = "   Image Scaler, allows you to zoom the image view in or out (5 - 500%."
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "   Capturing image files" Then
        frmHelp1.Visible = False
        frmHelp2.Visible = True
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
        frmHelp2.Label1.Caption = "   Move your mouse pointer over either the capture screen,rectangle or color controls at the bottom right."
        Image1.Visible = True
    ElseIf Combo1.Text = "       Capture a rectangle." Then
        frmHelp1.Visible = True
        frmHelp2.Visible = False
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/capturerectangle.jpg")
        frmHelp1.Label1.Caption = "   Clicking here hides this form and captures (by drawing a rectangle around chosen area)a rectangular area, which is then loaded into the Image Viewer."
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Capture fullscreen." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/capturescreen.jpg")
        frmHelp1.Label1.Caption = "   Clicking here hides this form and captures a full screenshot which is then loaded into the Image Viewer."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "   Color Controls" Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/main1.jpg")
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        frmHelp1.Label1.Caption = "   Move your mouse pointer over a control."
        Image1.Visible = True
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       H,S, & B Controls." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/main1.jpg")
        frmHelp1.Label1.Caption = "   Move your mouse pointer over an H,S or B option control."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = True
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       RGB Controls." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/main1.jpg")
        frmHelp1.Label1.Caption = "   Move your mouse pointer over a control in the RGB area."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Brightness Control." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/brightness.jpg")
        frmHelp1.Label1.Caption = "   Allows adjustment of the brightness of the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "   Capturing colors" Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/capturecolor.jpg")
        frmHelp1.Label1.Caption = "   Clicking here opens a small magnifier, which magnifys to the single pixel. It then captures the color you click on, either within or outside this program. This then becomes the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Web safe colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/webonly.jpg")
        frmHelp1.Label1.Caption = "   Sets the Color Picker to show web safe colors only."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Safe Pallette." Then
        frmHelp1.Visible = False
        frmHelp2.Visible = False
        frmHelp3.Visible = True
        frmInfo.Visible = False
        frmHelp3.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/safe.jpg")
        frmHelp3.Label1.Caption = "   Web safe color picker, clicking here loads web safe color choice as selected color."
        frmHelp4.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 805
        cmdStop.Visible = True
        Label2.Caption = "Safe Pallette"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipSafe.jpg")
     ElseIf Combo1.Text = "       Custom colors." Then
        frmHelp1.Visible = False
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = True
        frmInfo.Visible = False
        frmHelp4.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/custom1.jpg")
        frmHelp4.Label1.Caption = "   Custom colors, after you have selected or captured a color, click a box here. Then click the add button and this color will be saved to your custom pallette"
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Capture from screen." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/capturecolor.jpg")
        frmHelp1.Label1.Caption = "   Clicking here opens a small magnifier, which magnifys to the single pixel. It then captures the color you click on, either within or outside this program. This then becomes the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Using the Picker." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/main1.jpg")
        frmHelp1.Label1.Caption = "   Main color picker, clicking on a color here displays the various color codes for selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       System Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/syscolorcombo.jpg")
        frmHelp1.Label1.Caption = "   System color picker, use this to load codes for system colors."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 405
        cmdStop.Visible = True
        Label2.Caption = "System Colors"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipSystem.jpg")
    ElseIf Combo1.Text = "       CMYK Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/cyan.jpg")
        frmHelp1.Label1.Caption = "   Shows the percentage of cyan,magenta,yellow and black in the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 605
        cmdStop.Visible = True
        Label2.Caption = "CMYK Color"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipCMYK.jpg")
    ElseIf Combo1.Text = "       QB Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/qbpanel.jpg")
        frmHelp1.Label1.Caption = "   QB color panel, you may select a QuickBasic color, for use in the older programming language or Visual Basic."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       HTML Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/htmlcode.jpg")
        frmHelp1.Label1.Caption = "   Shows the HTML code of the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 755
        cmdStop.Visible = True
        Label2.Caption = "Safe Pallette"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipHTML.jpg")
    ElseIf Combo1.Text = "       VB Hex Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/vbhexcode.jpg")
        frmHelp1.Label1.Caption = "   Shows the Visual Basic Hex code of the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "       Decimal Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/decimalcode.jpg")
        frmHelp1.Label1.Caption = "   Shows the decimal code of the selected color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 805
        cmdStop.Visible = True
        Label2.Caption = "Decimal Code"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipDecimal.jpg")
    ElseIf Combo1.Text = "       Pantone Colors." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/pantone.jpg")
        frmHelp1.Label1.Caption = "   Use this combobox to select a Pantone color."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 755
        cmdStop.Visible = True
        Label2.Caption = "Pantone Colors"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipPantone.jpg")
    ElseIf Combo1.Text = "    Tray Control." Then
        frmHelp1.Visible = True
        frmHelp1.Picture1.Picture = LoadPicture(App.Path & "/Images/Help/hide.jpg")
        frmHelp1.Label1.Caption = "   Hides this program and places an icon in the system tray. Right clicking on the icon shows a popup menu."
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        frmInfo.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "    More info about color." Then
        frmInfo.Visible = True
        frmHelp1.Visible = False
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = False
        Picture1.Visible = False
        cmdStop.Visible = False
    ElseIf Combo1.Text = "    View Read Me." Then
        frmInfo.Visible = True
        frmHelp1.Visible = False
        frmHelp2.Visible = False
        frmHelp3.Visible = False
        frmHelp4.Visible = False
        Image1.Visible = False
        timerScroll.Enabled = True
        Picture1.Visible = True
        Info.Top = 4
        Info.Height = 805
        cmdStop.Visible = True
        Label2.Caption = "Read Me File"
        Info.Picture = LoadPicture(App.Path & "/Images/Help/tipRead.jpg")
    End If
    
    
End Sub


Private Sub Form_Load()
Image1.Visible = True
timerScroll.Enabled = False
Picture1.Visible = False
cmdStop.Visible = False
Combo1.AddItem "Using ColorWhiz"
Combo1.AddItem "   Controls Overview."
Combo1.AddItem "   Using image files"
Combo1.AddItem "       Load file."
Combo1.AddItem "       Open file."
Combo1.AddItem "       Save file as."
Combo1.AddItem "       Copy file."
Combo1.AddItem "       Print image file."
Combo1.AddItem "       Thumbnail Viewer."
Combo1.AddItem "       Zoom Contol."
Combo1.AddItem "   Capturing image files"
Combo1.AddItem "       Capture a rectangle."
Combo1.AddItem "       Capture fullscreen."
Combo1.AddItem "   Color Controls"
Combo1.AddItem "       H,S, & B Controls."
Combo1.AddItem "       RGB Controls."
Combo1.AddItem "       Brightness Control."
Combo1.AddItem "   Capturing colors"
Combo1.AddItem "       Web safe colors."
Combo1.AddItem "       Safe Pallette."
Combo1.AddItem "       Custom colors."
Combo1.AddItem "       Capture from screen."
Combo1.AddItem "       Using the Picker."
Combo1.AddItem "       System Colors."
Combo1.AddItem "       CMYK Colors."
Combo1.AddItem "       QB Colors."
Combo1.AddItem "       HTML Colors."
Combo1.AddItem "       VB Hex Colors."
Combo1.AddItem "       Decimal Colors."
Combo1.AddItem "       Pantone Colors."
Combo1.AddItem "    Tray Control."
Combo1.AddItem "    More info about color."
Combo1.AddItem "    View Read Me."
End Sub

Private Sub timerScroll_Timer()
DoEvents
Info.Top = Info.Top - 1
End Sub
