VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ColorWhiz 2.1"
   ClientHeight    =   5910
   ClientLeft      =   390
   ClientTop       =   2730
   ClientWidth     =   10260
   Icon            =   "frmChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmChart.frx":076A
   Picture         =   "frmChart.frx":0A74
   ScaleHeight     =   394
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   684
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   9
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   100
      Text            =   "0xC7C2C2"
      ToolTipText     =   "System Color Code"
      Top             =   4020
      Width           =   1470
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   8
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   99
      Text            =   "$0000FFFF"
      ToolTipText     =   "System Color Code"
      Top             =   3540
      Width           =   1470
   End
   Begin VB.HScrollBar hsbShade 
      Height          =   120
      LargeChange     =   16
      Left            =   2220
      Max             =   255
      TabIndex        =   89
      Top             =   5490
      Width           =   4275
   End
   Begin VB.CheckBox chkHex 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E3E3&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5940
      TabIndex        =   86
      ToolTipText     =   "View Hex Code"
      Top             =   4140
      Width           =   195
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6195
      MaxLength       =   3
      TabIndex        =   85
      Top             =   4950
      Width           =   315
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6195
      MaxLength       =   3
      TabIndex        =   84
      Top             =   4650
      Width           =   315
   End
   Begin VB.TextBox txtRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6195
      MaxLength       =   3
      TabIndex        =   83
      Top             =   4350
      Width           =   315
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   120
      Index           =   2
      LargeChange     =   16
      Left            =   2760
      Max             =   255
      TabIndex        =   82
      Top             =   5010
      Width           =   3405
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   120
      Index           =   1
      LargeChange     =   16
      Left            =   2760
      Max             =   255
      TabIndex        =   81
      Top             =   4710
      Width           =   3405
   End
   Begin VB.HScrollBar hsbRGB 
      Height          =   120
      Index           =   0
      LargeChange     =   16
      Left            =   2760
      Max             =   255
      TabIndex        =   80
      Top             =   4410
      Width           =   3405
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   7
      Left            =   52
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "&H8000000F&"
      ToolTipText     =   "System Color Code"
      Top             =   4935
      Width           =   1890
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   6
      Left            =   52
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "vbButtonFace"
      ToolTipText     =   "System Color Name"
      Top             =   4635
      Width           =   1890
   End
   Begin VB.ComboBox cboSysColors 
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   52
      Style           =   2  'Dropdown List
      TabIndex        =   72
      ToolTipText     =   "System Colors"
      Top             =   4320
      Width           =   1890
   End
   Begin VB.VScrollBar IconScroll 
      Height          =   3495
      Left            =   7350
      Max             =   0
      TabIndex        =   64
      Top             =   120
      Width           =   155
   End
   Begin VB.DirListBox Folder 
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   1890
      Left            =   60
      TabIndex        =   62
      ToolTipText     =   "Folderlist"
      Top             =   360
      Width           =   1860
   End
   Begin VB.DriveListBox Drive 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   60
      TabIndex        =   61
      ToolTipText     =   "Drivelist"
      Top             =   45
      Width           =   1860
   End
   Begin VB.FileListBox FileList 
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   1845
      Left            =   45
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   60
      ToolTipText     =   "Filelist"
      Top             =   2235
      Width           =   1875
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A87856&
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   6450
      ScaleHeight     =   3450
      ScaleWidth      =   795
      TabIndex        =   52
      ToolTipText     =   "Thumbnails"
      Top             =   120
      Width           =   825
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   65
         Top             =   660
         Width           =   675
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   0
         Left            =   50
         Stretch         =   -1  'True
         Top             =   15
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   1
         Left            =   50
         Stretch         =   -1  'True
         Top             =   900
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   2
         Left            =   50
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   3
         Left            =   50
         Stretch         =   -1  'True
         Top             =   2700
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   4
         Left            =   50
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   5
         Left            =   50
         Stretch         =   -1  'True
         Top             =   4500
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   6
         Left            =   50
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   59
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   58
         Top             =   2460
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   57
         Top             =   3360
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   56
         Top             =   4260
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   55
         Top             =   5160
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   30
         TabIndex        =   54
         Top             =   6060
         Width           =   675
      End
      Begin VB.Shape Pic2 
         BackColor       =   &H80000001&
         BorderWidth     =   5
         FillColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image picIcon 
         Height          =   615
         Index           =   7
         Left            =   50
         Stretch         =   -1  'True
         Top             =   6300
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   30
         TabIndex        =   53
         Top             =   6960
         Width           =   675
      End
   End
   Begin VB.HScrollBar HMove 
      Height          =   155
      Left            =   2055
      Max             =   0
      TabIndex        =   51
      Top             =   3675
      Width           =   4125
   End
   Begin VB.HScrollBar ScaleBar 
      Enabled         =   0   'False
      Height          =   140
      Left            =   2580
      Max             =   500
      Min             =   5
      TabIndex        =   45
      Top             =   3870
      Value           =   100
      Width           =   3605
   End
   Begin VB.VScrollBar VMove 
      Height          =   3495
      Left            =   6255
      Max             =   0
      TabIndex        =   44
      Top             =   120
      Width           =   155
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   4980
      Top             =   5760
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmChart.frx":B748
      OLEDBString     =   $"frmChart.frx":B7D6
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cboPantone 
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6855
      TabIndex        =   41
      Text            =   " Pantone Colors"
      ToolTipText     =   "Select A Pantone Color"
      Top             =   4545
      Width           =   3210
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   2100
      ScaleHeight     =   3300
      ScaleWidth      =   4140
      TabIndex        =   38
      Top             =   165
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3360
      Left            =   2100
      ScaleHeight     =   3300
      ScaleWidth      =   4140
      TabIndex        =   39
      Top             =   90
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3480
      Left            =   2040
      ScaleHeight     =   3420
      ScaleWidth      =   4275
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Timer ReDrawTimer 
      Interval        =   1
      Left            =   5940
      Top             =   5760
   End
   Begin VB.CheckBox chkCapture 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6060
      TabIndex        =   36
      Top             =   5760
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   5760
   End
   Begin VB.Frame fraQBColors 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   6420
      TabIndex        =   31
      Top             =   360
      Width           =   1050
      Begin VB.TextBox txtColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "vbConst"
         Top             =   3350
         Width           =   1040
      End
      Begin VB.PictureBox picQBColors 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3360
         Left            =   0
         ScaleHeight     =   222
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   90
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.CheckBox chkMsg 
      Height          =   195
      Left            =   6060
      TabIndex        =   30
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "C++ Color"
      ToolTipText     =   "QB Color Code"
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Height          =   285
      Index           =   2
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "RGB(255, 255, 255)"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   3
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "16777215"
      ToolTipText     =   "Decimal Code"
      Top             =   3060
      Width           =   1455
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "#FFFF00"
      ToolTipText     =   "HTML Code"
      Top             =   2100
      Width           =   1455
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "&H00FFFF&"
      ToolTipText     =   "VB Hex Code"
      Top             =   2580
      Width           =   1455
   End
   Begin VB.ComboBox cmbPreset 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8055
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   390
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E1E1E1&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   7680
      TabIndex        =   6
      Top             =   120
      Width           =   2490
      Begin VB.CheckBox chkWeb 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1E1E1&
         Caption         =   "Only web colors."
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   10
         TabIndex        =   93
         ToolTipText     =   "Only Web Safe Colors"
         Top             =   990
         Width           =   1530
      End
      Begin VB.TextBox txtK 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   20
         Text            =   "100"
         ToolTipText     =   "Black"
         Top             =   900
         Width           =   405
      End
      Begin VB.TextBox txtYellow 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   19
         Text            =   "100"
         ToolTipText     =   "Yellow"
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtMagenta 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   18
         Text            =   "100"
         ToolTipText     =   "Magenta"
         Top             =   300
         Width           =   405
      End
      Begin VB.TextBox txtCyan 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   "Cyan"
         Top             =   0
         Width           =   405
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Blue"
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Green"
         Top             =   300
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "255"
         ToolTipText     =   "Red"
         Top             =   0
         Width           =   405
      End
      Begin VB.TextBox txtH 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   470
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Hue"
         Top             =   0
         Width           =   405
      End
      Begin VB.TextBox txtS 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   470
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   4
         Text            =   "100"
         ToolTipText     =   "Saturation"
         Top             =   300
         Width           =   405
      End
      Begin VB.TextBox txtB 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2F7F8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   470
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   5
         Text            =   "100"
         ToolTipText     =   "Brightness"
         Top             =   600
         Width           =   405
      End
      Begin VB.OptionButton optH 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1E1E1&
         Caption         =   "H:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Hue"
         Top             =   0
         Width           =   525
      End
      Begin VB.OptionButton optS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1E1E1&
         Caption         =   "S:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Saturation"
         Top             =   300
         Width           =   555
      End
      Begin VB.OptionButton optB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1E1E1&
         Caption         =   "B:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Brightness"
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " K:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1815
         TabIndex        =   24
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Y:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1815
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " M:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1815
         TabIndex        =   22
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " C:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1815
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblR 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " R:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1130
         TabIndex        =   15
         ToolTipText     =   "Red"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblG 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " G:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1130
         TabIndex        =   14
         ToolTipText     =   "Green"
         Top             =   300
         Width           =   375
      End
      Begin VB.Label lblB 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " B:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1130
         TabIndex        =   13
         ToolTipText     =   "Blue"
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   930
         TabIndex        =   12
         Top             =   300
         Width           =   165
      End
      Begin VB.Label lblBB 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   930
         TabIndex        =   11
         Top             =   600
         Width           =   195
      End
      Begin VB.Label lblH 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   930
         TabIndex        =   10
         Top             =   -20
         Width           =   90
      End
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   6000
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FCE4D2&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2055
      ScaleHeight     =   3495
      ScaleWidth      =   4125
      TabIndex        =   46
      Top             =   120
      Width           =   4125
      Begin VB.PictureBox NewImage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00A87856&
         ForeColor       =   &H80000008&
         Height          =   3490
         Left            =   0
         MousePointer    =   5  'Size
         ScaleHeight     =   3440.228
         ScaleMode       =   0  'User
         ScaleWidth      =   4095
         TabIndex        =   47
         Top             =   0
         Width           =   4125
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00A87856&
         ForeColor       =   &H80000008&
         Height          =   4175
         Left            =   0
         ScaleHeight     =   4140
         ScaleWidth      =   4095
         TabIndex        =   48
         Top             =   0
         Width           =   4125
         Begin VB.PictureBox Sample 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00A87856&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5595
            Left            =   0
            ScaleHeight     =   5615.217
            ScaleMode       =   0  'User
            ScaleWidth      =   4125
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   4125
            Begin VB.PictureBox picRectangle 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   855
               Left            =   4020
               ScaleHeight     =   57
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   97
               TabIndex        =   50
               Top             =   -60
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   4980
               Top             =   840
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
      End
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Java Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   7740
      TabIndex        =   102
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Delphi Hex Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   7740
      TabIndex        =   101
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Image cmdHide 
      Height          =   450
      Left            =   9300
      Picture         =   "frmChart.frx":B864
      ToolTipText     =   "Hide Me"
      Top             =   5340
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image cmdHelp 
      Height          =   555
      Left            =   8700
      Picture         =   "frmChart.frx":BB7E
      ToolTipText     =   "Help/About"
      Top             =   5280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image cmdPreset 
      Height          =   615
      Left            =   8040
      Picture         =   "frmChart.frx":BFAC
      ToolTipText     =   "Pallette"
      Top             =   5280
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image cmdRect 
      Height          =   600
      Left            =   7020
      Picture         =   "frmChart.frx":C441
      ToolTipText     =   "Capture Rectangular Area"
      Top             =   5280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgHelp 
      Height          =   525
      Left            =   8760
      Picture         =   "frmChart.frx":C836
      Top             =   5280
      Width           =   465
   End
   Begin VB.Image cmdPrint 
      Height          =   600
      Left            =   1620
      Picture         =   "frmChart.frx":CC23
      ToolTipText     =   "Print"
      Top             =   5295
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image cmdCopy 
      Height          =   555
      Left            =   1095
      Picture         =   "frmChart.frx":D082
      ToolTipText     =   "Copy"
      Top             =   5340
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image cmdSave 
      Height          =   585
      Left            =   600
      Picture         =   "frmChart.frx":D4BD
      ToolTipText     =   "Save File"
      Top             =   5325
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image cmdLoad 
      Height          =   525
      Left            =   0
      Picture         =   "frmChart.frx":D902
      ToolTipText     =   "Open File"
      Top             =   5400
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgPrint 
      Height          =   525
      Left            =   1620
      Picture         =   "frmChart.frx":DD1B
      Top             =   5400
      Width           =   495
   End
   Begin VB.Image imgCopy 
      Height          =   525
      Left            =   1080
      Picture         =   "frmChart.frx":E136
      Top             =   5400
      Width           =   495
   End
   Begin VB.Image imgSave 
      Height          =   525
      Left            =   540
      Picture         =   "frmChart.frx":E530
      Top             =   5400
      Width           =   525
   End
   Begin VB.Image imageLoad 
      Height          =   525
      Left            =   0
      Picture         =   "frmChart.frx":E944
      Top             =   5400
      Width           =   555
   End
   Begin VB.Image cmdPicker 
      Height          =   600
      Left            =   7500
      Picture         =   "frmChart.frx":ED5F
      ToolTipText     =   "Capture Color"
      Top             =   5280
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image cmdExit 
      Height          =   540
      Left            =   9720
      Picture         =   "frmChart.frx":F1DD
      ToolTipText     =   "Exit"
      Top             =   5280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image cmdCapture 
      Height          =   600
      Left            =   6480
      Picture         =   "frmChart.frx":F60C
      ToolTipText     =   "Capture Full Screen"
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgPreset 
      Height          =   525
      Left            =   8100
      Picture         =   "frmChart.frx":FA4E
      Top             =   5325
      Width           =   600
   End
   Begin VB.Image imgADD 
      Height          =   360
      Left            =   8955
      Picture         =   "frmChart.frx":FE95
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image cmdADD 
      Height          =   360
      Left            =   8955
      Picture         =   "frmChart.frx":102E7
      ToolTipText     =   "Add to Custom Pallette"
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblADDColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   8070
      TabIndex        =   98
      Top             =   810
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Safe Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8340
      TabIndex        =   97
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scale:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2040
      TabIndex        =   96
      Top             =   3840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   680
      X2              =   680
      Y1              =   4
      Y2              =   92
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   508
      X2              =   508
      Y1              =   4
      Y2              =   92
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   508
      X2              =   680
      Y1              =   92
      Y2              =   92
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   506
      X2              =   680
      Y1              =   4
      Y2              =   4
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   6225
      TabIndex        =   95
      Top             =   4140
      Width           =   345
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pantone Colors:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   6840
      TabIndex        =   94
      Top             =   4320
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8340
      TabIndex        =   92
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblLighter 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lighter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5880
      TabIndex        =   88
      Top             =   5280
      Width           =   600
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Darker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2220
      TabIndex        =   87
      Top             =   5280
      Width           =   585
   End
   Begin VB.Label lblBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   2700
      TabIndex        =   79
      Top             =   4980
      Width           =   3795
   End
   Begin VB.Label lblBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   2700
      TabIndex        =   78
      Top             =   4680
      Width           =   3795
   End
   Begin VB.Label lblBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   2700
      TabIndex        =   77
      Top             =   4380
      Width           =   3795
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " RGB Color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   2220
      TabIndex        =   76
      Top             =   4140
      Width           =   1020
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "System Colors:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   75
      Top             =   4140
      Width           =   1260
   End
   Begin VB.Label lblSpyColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9300
      TabIndex        =   70
      ToolTipText     =   "Mouse is on Color"
      Top             =   2580
      Width           =   840
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pointer"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9480
      TabIndex        =   71
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9300
      TabIndex        =   68
      ToolTipText     =   "Old Color"
      Top             =   2100
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Color"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9405
      TabIndex        =   69
      Top             =   1920
      Width           =   645
   End
   Begin VB.Label lblSelColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9300
      TabIndex        =   67
      ToolTipText     =   "New Color"
      Top             =   1620
      Width           =   840
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Color"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   9345
      TabIndex        =   66
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6720
      TabIndex        =   63
      Top             =   3600
      Width           =   405
   End
   Begin VB.Label ScaleMeter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6210
      TabIndex        =   43
      Top             =   3840
      Width           =   390
   End
   Begin VB.Label lblPantone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D2F7F8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Select a Pantone Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6855
      TabIndex        =   42
      ToolTipText     =   "Pantone Color"
      Top             =   4875
      Width           =   3225
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "C++ Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   7680
      TabIndex        =   35
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   132
      X2              =   504
      Y1              =   272
      Y2              =   272
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   132
      X2              =   504
      Y1              =   4
      Y2              =   4
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   132
      X2              =   132
      Y1              =   3
      Y2              =   272
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   504
      X2              =   504
      Y1              =   4
      Y2              =   272
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "QB Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6480
      TabIndex        =   40
      Top             =   120
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   3840
      Picture         =   "frmChart.frx":106F5
      Top             =   3660
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3840
      Picture         =   "frmChart.frx":10AEA
      Top             =   3660
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " Decimal Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   7680
      TabIndex        =   34
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " HTML Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   7680
      TabIndex        =   33
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label lblCap 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " VB Hex Color:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   7680
      TabIndex        =   32
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image imgRect 
      Height          =   525
      Left            =   7080
      Picture         =   "frmChart.frx":10F2F
      Top             =   5340
      Width           =   435
   End
   Begin VB.Image imgPicker 
      Height          =   525
      Left            =   7560
      Picture         =   "frmChart.frx":112FB
      Top             =   5340
      Width           =   510
   End
   Begin VB.Image imgExit 
      Height          =   435
      Left            =   9780
      Picture         =   "frmChart.frx":11740
      Top             =   5340
      Width           =   420
   End
   Begin VB.Image imgCapture 
      Height          =   525
      Left            =   6540
      Picture         =   "frmChart.frx":11B15
      Top             =   5340
      Width           =   480
   End
   Begin VB.Image imgHide 
      Height          =   390
      Left            =   9300
      Picture         =   "frmChart.frx":11F0B
      Top             =   5370
      Width           =   390
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "Capture"
      Visible         =   0   'False
      Begin VB.Menu mnuScreen 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuRectangle 
         Caption         =   "Rectangle"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About ColorWhiz"
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "Help"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show Me"
      End
      Begin VB.Menu mnuTrayCapture 
         Caption         =   "Capture"
         Begin VB.Menu mnuTrayScreen 
            Caption         =   "Full Screen"
         End
         Begin VB.Menu mnuTrayRectangle 
            Caption         =   "Rectangle"
         End
         Begin VB.Menu mnuTrayColor 
            Caption         =   "Color"
         End
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rs As New ADODB.Recordset 'holds variable for recordset connection
Private mbCrop As Boolean
Private mbDown As Boolean
Private nOldX As Integer
Private nOldY As Integer
Private mbNoChange      As Boolean
Private mbCapture       As Boolean
Private miCurQBIdx      As Integer
Private mlCurColor      As Long
Private msaSysColors()  As String
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "User32" () As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const mcnTITLE_BAR_HEIGHT As Integer = 400
Private Const SRCCOPY = &HCC0020
Private Type PointAPI
    X   As Long
    Y   As Long
End Type

Enum sMode
    Picker = 0
    About = 1
    Custom = 2
    SafeColor = 3
End Enum

Dim MainBoxHit  As Boolean
Dim SelectBoxHit As Boolean
Public Mode As sMode
Public UserMode As Integer 'usercontrol Mode variable
Dim HueEntering As Boolean
Dim SaturationEntering As Boolean
Dim BrightnessEntering As Boolean
Dim OldHue As Integer, OldSaturation As Integer, OldBrightness As Integer
Dim ChangeByInput As Boolean
Dim Loaded As Boolean
Dim LcButtonRect As RECT
Dim LcSmpRect As RECT
Dim lcButtonPressed As Boolean
Dim Tempstring(1 To 3000) As Variant
Dim ipicHeight As Integer
Dim ipicWidth As Integer
Dim lYOffset As Integer
Dim iColorCur As Single
Dim iColorStep As Single
Dim NumLines As Integer
Dim lX As Long
Dim lY As Long
Dim strRead As String
Dim IconStart As Integer
Dim DragX As Double
Dim DragY As Double
Dim Dragnow As Boolean
Dim WWW As Double
Dim HHH As Double
Dim FileExt As String
Dim Kill_or_Not As Boolean
Dim ClipboardPath As String
Dim DefaultDrive As String
Dim DefaultPath As String
Dim FilePath As String
Dim Picture_Shown As Boolean
Dim NewImage_Width As Double
Dim NewImage_Height As Double
Dim Path$
Dim Filename$
Dim I, b As Integer
Dim r$
Dim j%
Private Sub cboPantone_Click()
'************************************
    'This sub finds items selected in combo
    '************************************
 If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If

    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        Picture1.Visible = False
        Picture2.Visible = False
        IconScroll.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If
    
    Me.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    
    
        cmbPreset.Visible = False
        chkCapture.Enabled = True
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
        
    Select Case True
        Case optH
            optH_Click
        Case optS
            optS_Click
        Case optB
            optB_Click
    End Select
    
    Frame2.Visible = True
    imgADD.Visible = False
    lblADDColor.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Mode = Picker
    
    On Error Resume Next
    
    rs.MoveFirst
    rs.Find "PANTONENAM= '" & cboPantone & "'"
    Call LoadFields
End Sub
Private Sub LoadFields()
    '************************************
    'Load fields with values in rs
    '************************************
    
    On Error Resume Next
    Text1.Text = "" & rs("R")
    Text2.Text = "" & rs("G")
    Text3.Text = "" & rs("B")
    SetColors
    lblPantone.Caption = "" & rs("PantoneNam")

End Sub

Private Sub cmdCopy_Click()
    
    If NewImage.Picture <> 0 Then
        Clipboard.Clear
        DoEvents
        Clipboard.SetData NewImage.Picture
        DoEvents
        MsgBox "Image saved to Windows clipboard.  Use Paste or CTL+V in another application, such as Word, to paste image from clipboard.", vbInformation, "Copy Image to Clipboard"
    Else
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Please capture or load an image before copying to clipboard.", vbInformation, "Nothing To Copy"
    End If

End Sub

Private Sub cmdLoad_Click()

On Error Resume Next
    frmChart.Cls
    Mode = SafeColor
    imgPreset.Visible = False
    cmdPreset.Visible = False
    IconStart = 0
    Update_Icons
    Picture1.Visible = True
    Picture2.Visible = True
    lblInfo.Visible = True
    IconScroll.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    cd2.Filter = "Image Files (*.*)|*.*"
    cd2.ShowOpen
    Filename$ = cd2.Filename
    Picture_Shown = True
    FilePath = cd2.InitDir + r$ + cd2.Filename
    NewImage.Visible = True
    Sample.Picture = LoadPicture(Filename$)
    NewImage.Picture = LoadPicture(Filename$)
    NewImage_Width = NewImage.Width
    NewImage_Height = NewImage.Height
    ScaleBar.Enabled = True
    ScaleBar.Value = 100
    ScaleBar.Max = 500
        If NewImage_Width * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Width) * 100
        If NewImage_Height * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Height) * 100
    NewImage.Cls
    Picture_Update

End Sub

Private Sub cmdPicker_Click()

chkCapture.Value = 1

End Sub

Private Sub cmdPrint_Click()
If (Picture2.Visible = True) And (NewImage.Picture <> 0) Then
        frmPrint.PrintBitmap NewImage.Picture
    Else
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Please capture or load an image before printing.", vbInformation, "Nothing To Print"
End If

End Sub

Private Sub cmdRect_Click()
    
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
    End If
    
    If CaptureDesktop Then
        '--- Show the Form where all the Work will take place
        DoEvents
        
        frmRectangle.ShowPicture NewImage.Picture
    End If
    
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")

End Sub

Private Sub cmdSave_Click()

    If NewImage.Picture = 0 Or Picture2.Visible = False Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Please capture or load an image before saving.", vbInformation, "Nothing To Save"
    Else
        '--- Set the Filters
        With CD1
        
            .Filter = "Bmp (*.bmp)|*.bmp|Gif (*.gif)|*.gif|Jpeg (*.jpeg;*.jpg;*.jpe;*.jfif)|*.jpeg;*.jpg;*.jpe;*.jfif|TIFF (*.tiff;*.tif)|*.tiff;*.tif|Png (*.png)|*.png"
            '--- Specify default filter
            .FilterIndex = 2
            '--- Hide the "Open as read only" checkbox when saving.
            .Flags = cdlOFNHideReadOnly
            '--- Show the Open Dialog
            .ShowSave
            '--- If Canceled is Pressed
            If .Filename = "" Then Exit Sub
            '--- Save the Image
            SavePicture NewImage.Image, .Filename
        End With
    End If
    

End Sub

Private Sub cmdHelp_Click()

PopupMenu mnuHelp

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Image1.Visible = False
    Image2.Visible = True
    
End Sub
Public Sub LoadMode(ByVal lMode As Integer)
    
    Dim cl As Long
    Dim mc As Long
    Dim hs As HSB, hs1 As HSB
    Dim Red As Integer, Green As Integer, Blue As Integer
    pMode = lMode
    cl = lblSelColor.BackColor ' m_Color
    GetRGB cl, Red, Green, Blue
    frmChart.Cls
    'PrintLastColor
    Select Case lMode
    Case 0
        LoadMainHue
        hs = RGBtoHSB(lblSelColor.BackColor)
        SelectedMainPos = MainBox.Top + (360 - hs.Hue) * 255 / 360
        Call DrawSlider(SelectedMainPos)
        mc = GetPixel(Me.hDC, MainBox.Left + 5, SelectedMainPos)
        GetRGB mc, Red, Green, Blue
        SelectedPos.X = SelectBox.Left + hs.Saturation * 255 / 100
        SelectedPos.Y = SelectBox.Bottom - hs.Brightness * 255 / 100
        
        LoadVariantsHue Red, Green, Blue
        DrawPicker
    Case 1
        hs = RGBtoHSB(lblSelColor.BackColor)
        hs1 = hs
        LoadVariantsSaturation hs.Saturation / 100
        SelectedPos.X = SelectBox.Left + hs.Hue * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - (hs.Brightness) * 255 / 100
        SelectedMainPos = MainBox.Top + (100 - hs.Saturation) * 255 / 100
        Call DrawSlider(SelectedMainPos)
        hs1.Saturation = 100
        hs1.Brightness = 100
        HSBtoRGB hs1, Red, Green, Blue
        DrawPicker
        LoadMainSaturation Me.hDC, Red, Green, Blue, hs.Brightness / 100
    Case 2
        hs = RGBtoHSB(lblSelColor.BackColor)
        hs1 = hs
        LoadVariantsBrightness hs.Brightness / 100
        SelectedPos.X = SelectBox.Left + hs.Hue * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - hs.Saturation * 255 / 100
        SelectedMainPos = MainBox.Top + 255 - (hs.Brightness * 255 / 100)
        Call DrawSlider(SelectedMainPos)
        'hs1.Saturation = 100
        hs1.Brightness = 100
        HSBtoRGB hs1, Red, Green, Blue
        LoadMainBrightness Me.hDC, Red, Green, Blue
        DrawPicker
    End Select
    mc = lblSelColor.BackColor  'm_Color
    GetRGB mc, Red, Green, Blue
    Text1.Text = Red
    Text2.Text = Green
    Text3.Text = Blue
    txtH.Text = CInt(hs.Hue)
    txtS.Text = CInt(hs.Saturation)
    txtB.Text = CInt(hs.Brightness)
    Me.PSet (-100, -100)
    
End Sub

Private Sub chkWeb_Click()
    
    m_WebColors = IIf(chkWeb.Value = vbChecked, True, False)
        If Loaded = False Then Exit Sub
    ReloadColors

End Sub

Private Sub ReloadColors()

If Mode = Picker Then
    Dim cl As Long
    Dim Red As Integer, Green As Integer, Blue As Integer
    
    DrawPicker
    DrawSlider SelectedMainPos
    Dim hs As HSB
    Select Case True
    Case optH.Value
        SelectedMainPos = MainBox.Top + (360 - Val(txtH.Text)) * 255 / 360
        SelectedPos.X = SelectBox.Left + Val(txtS.Text) * 255 / 100
        SelectedPos.Y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
        LoadMainHue
        cl = GetPixel(frmChart.hDC, MainBox.Left + 5, SelectedMainPos)
        GetRGB cl, Red, Green, Blue
        LoadVariantsHue Red, Green, Blue
        DrawPicker
    Case optS.Value
        SelectedMainPos = MainBox.Top + (100 - Val(txtS.Text)) * 255 / 100
        SelectedPos.X = SelectBox.Left + Val(txtH.Text) * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
        LoadVariantsSaturation Val(txtS.Text) / 100
        hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
        HSBtoRGB hs, Red, Green, Blue
        LoadMainSaturation frmChart.hDC, Red, Green, Blue, Val(txtB.Text) / 100
        DrawPicker
    Case optB.Value
        SelectedMainPos = MainBox.Top + (100 - Val(txtB.Text)) * 255 / 100
        SelectedPos.X = SelectBox.Left + Val(txtH.Text) * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - Val(txtS.Text) * 255 / 100
        LoadVariantsBrightness Val(txtB.Text) / 100
        hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
        HSBtoRGB hs, Red, Green, Blue
        LoadMainBrightness frmChart.hDC, Red, Green, Blue
        DrawPicker
    End Select
    
    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    'Java
    txtColor(9).Text = "0x" & StrRev(sHex)
    txtColor(5).Text = "&H" & sHex & "&"
    'Delphi
    txtColor(8).Text = "$00" & sHex
    'CPlus
    txtColor(1).Text = "0x00" & sHex
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
    DrawSlider SelectedMainPos
    
End If

End Sub
Private Sub cmbPreset_Click()
    
    frmChart.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim cl As Long
    Dim HexV As String
    On Error GoTo r:
    
    Select Case cmbPreset.ListIndex
    Case 0
        If picOut.Visible = True Then
            sndPlay "bloop_x", SoundOps.SND_ASYNC
            picBuffer.Visible = False
            picBackBuffer.Visible = False
            picOut.Visible = False
            Image1.Visible = False
            Image2.Visible = False
        End If
    
        If Picture2.Visible = True Then
            sndPlay "bloop_x", SoundOps.SND_ASYNC
            fraQBColors.Visible = True
            frmChart.Label8.Visible = True
            IconScroll.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            lblInfo.Visible = False
            VMove.Visible = False
            HMove.Visible = False
            Label13.Visible = False
            ScaleBar.Visible = False
            ScaleMeter.Visible = False
        End If
    
        pMode = 3
        Me.DrawStyle = 0
        Me.Picture = LoadPicture(App.Path & "/Images/main2.jpg")
        LoadCustomColors
        imgADD.Visible = True
        lblADDColor.Visible = True
        Label12.Visible = True
        Label14.Visible = False
        Mode = Custom
        imgPreset.Visible = True
        cmdPreset.Visible = False
        cl = lblSelColor.BackColor
        GetRGB cl, r, g, b
        GetHexVal r, g, b, HexV

    Case 1
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
    End If
    
    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        IconScroll.Visible = False
        Picture1.Visible = False
        Picture2.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
    End If
    
    pMode = 4
    Me.DrawStyle = 0
    Me.Picture = LoadPicture(App.Path & "/Images/main4.jpg")
    LoadSafePalette
    imgADD.Visible = False
    lblADDColor.Visible = False
    Label12.Visible = False
    Label14.Visible = True
    Mode = SafeColor
    imgPreset.Visible = True
    cmdPreset.Visible = False
    cl = lblSelColor.BackColor
    GetRGB cl, r, g, b
    GetHexVal r, g, b, HexV
    
    End Select
    Me.PSet (-100, -100)
    Exit Sub
r:
    Exit Sub

End Sub

Private Sub cmdADD_Click()
    
    svdColor(cPaletteIndex) = lblADDColor.BackColor
    DrawSafePicker cPaletteIndex, True 'Erase
    frmChart.DrawWidth = 1
    frmChart.DrawMode = 13
    frmChart.FillStyle = 0
    frmChart.ForeColor = 0
    frmChart.FillColor = svdColor(cPaletteIndex)
    Rectangle frmChart.hDC, Preset(cPaletteIndex).Left, Preset(cPaletteIndex).Top, Preset(cPaletteIndex).Right, Preset(cPaletteIndex).Bottom
    SaveCustomColors
    DrawSafePicker cPaletteIndex, False
    lblSelColor.BackColor = svdColor(cPaletteIndex)
    frmChart.PSet (-100, -100)

End Sub

Private Sub cmdCapture_Click()

'--- This is where we capture the FULL screen
'--- when the User chooses to capture the FULL Screen
    frmChart.Cls
    '--- Capture the Screen
        If picOut.Visible = True Then
            sndPlay "bloop_x", SoundOps.SND_ASYNC
            picBuffer.Visible = False
            picBackBuffer.Visible = False
            picOut.Visible = False
            Image1.Visible = False
            Image2.Visible = False
        End If
    
    CaptureDesktop
    ScaleBar.Enabled = True
    Mode = Custom
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    IconStart = 0
    Update_Icons
    Picture1.Visible = True
    Picture2.Visible = True
    IconScroll.Visible = True
    lblInfo.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
    Me.Top = 2400
    Me.Show
    
End Sub

Private Sub cmdExit_Click()
    
    sndPlay "erase", SoundOps.SND_ASYNC
    Dim Last As Integer
    Dim I As Integer
    On Error Resume Next
    Last = FreeFile()
    Open App.Path & "/lastcolor.cps" For Output As #Last
    Print #Last, lblSelColor.BackColor
    Close #Last
    Unload mdi_Help
    Unload frmMagnify
    Unload frmRectangle
    Unload frmPrint
    Unload Me

End Sub

Sub cmdPreset_Click()
    
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If

    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        Picture1.Visible = False
        Picture2.Visible = False
        IconScroll.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If
    
    Me.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    
    If Mode = Picker Then
        Frame2.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main2.jpg")
        Mode = Custom
        chkCapture.Enabled = False
        Label12.Visible = True
        Label14.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        cmbPreset.Visible = True
        LoadCustomColors
        cmbPreset.ListIndex = 1
        lblADDColor.BackColor = lblSelColor.BackColor
        GetRGB lblSelColor.BackColor, r, g, b
        GetHexVal r, g, b, HexV

    Else
        cmbPreset.Visible = False
        chkCapture.Enabled = True
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
        
    Select Case True
        Case optH
            optH_Click
        Case optS
            optS_Click
        Case optB
            optB_Click
    End Select
    
    Frame2.Visible = True
    imgADD.Visible = False
    lblADDColor.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Mode = Picker
    
    End If

End Sub

Function GetSafeColor(Index As Integer, r As Integer, g As Integer, b As Integer, Optional HexVal As String) As Long

    Dim I As Long, j As Long, k As Long
    Dim count As Integer
    Dim strR As String, strG As String, strB As String

    For I = 0 To &HFF Step &H33
        For j = 0 To &HFF Step &H33
            For k = 0 To &HFF Step &H33
            count = count + 1
                If count = Index Then
                    r = I: g = j: b = k
                    GetSafeColor = RGB(I, j, k)
                    GetHexVal r, g, b, HexVal
                Exit Function
            End If
        Next k
    Next j
Next I

End Function

Sub GetHexVal(Red As Integer, Green As Integer, Blue As Integer, StrHex As String)
    
    Dim strR As String, strG As String, strB As String
    strR = Trim(Hex(Red))
        If Len(strR) = 1 Then strR = "0" & strR
    strG = Trim(Hex(Green))
        If Len(strG) = 1 Then strR = "0" & strR
    strB = Trim(Hex(Blue))
        If Len(strB) = 1 Then strR = "0" & strR
        StrHex = strR & strG & strB

End Sub

Private Sub Form_Load()

    Dim iIdx As Integer
    Dim sMsg As String
    Dim sql As String
    DefaultDrive = Drive
    DefaultPath = Folder.Path
    IconStart = 0
    Update_Icons
    Picture1.Visible = False
    Picture2.Visible = False
    lblInfo.Visible = False
    IconScroll.Visible = False
    VMove.Visible = False
    HMove.Visible = False
    Label13.Visible = False
    ScaleBar.Visible = False
    ScaleMeter.Visible = False
        
        If App.PrevInstance Then
            sndPlay "BOING", SoundOps.SND_ASYNC
            MsgBox "ColorWhiz is Already Running" _
                & vbCrLf & vbCrLf & "Please check the System Tray or Your Taskbar", _
            vbInformation, "Previous Instance"
        End
        End If
     
    
   '************************************
    'Establish connection to data source
    '************************************
    sql = "SELECT * FROM tblPantoneConvTable"
    rs.Open sql, gConn, adOpenKeyset, adLockOptimistic
    
    Call LoadPantoneCbo
    Call LoadFields
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    DrawQBButtons
    miCurQBIdx = -1
    GetSystemColors
    
    For iIdx = 0 To txtColor.UBound
        txtColor(iIdx).ToolTipText = "Right-click to copy this color code"
    Next
    
    picQBColors_MouseDown 1, 0, picQBColors.ScaleWidth - 5, picQBColors.ScaleHeight - 5
    
    If UCase$(Command$) = "/H" Then
        Me.Hide
    End If
    ' Set these Parameters on Basis of where  should Select Box and  Main Box Should appear
    Dim OldP As PointAPI
    ReDim Preset(1 To 224)  'RECT structure
    Preset(1).Left = 138
    Preset(1).Top = 10
    Preset(1).Right = 153
    Preset(1).Bottom = 25
    
    With LcButtonRect
        .Right = Me.ScaleWidth - 8
        .Left = .Right - 55
        .Top = 25
        .Bottom = .Top + 20
    End With
    
    With LcSmpRect
        .Left = LcButtonRect.Left + 4
        .Top = LcButtonRect.Top + 4
        .Bottom = LcButtonRect.Bottom - 4
        .Right = .Left + .Bottom - .Top + 6
    End With
    
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Mode = Picker   'Initialize Mode as  normal Picker
    
    With lpBI.bmiHeader
        .biBitCount = 24
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(lpBI.bmiHeader)
    End With
    
    '// Setting the position of RECTS for Safe Colors
    For I = 2 To 224
        If I Mod 16 = 0 Then
            Preset(I).Top = Preset(I - 1).Top
            Preset(I).Left = Preset(I - 1).Right + 3
            Preset(I).Bottom = Preset(I).Top + 15
            Preset(I).Right = Preset(I).Left + 15
            If I = 224 Then GoTo Jump
            I = I + 1
            Preset(I).Top = Preset(I - 1).Bottom + 3
            Preset(I).Left = Preset(1).Left
        Else
            Preset(I).Top = Preset(I - 1).Top
            Preset(I).Left = Preset(I - 1).Right + 3
        End If
Jump:
        Preset(I).Bottom = Preset(I).Top + 15
        Preset(I).Right = Preset(I).Left + 15
    Next I

    k = 255
    'Show
    SelectBox.Left = 138
    SelectBox.Top = 10
    SelectBox.Right = SelectBox.Left + k
    SelectBox.Bottom = SelectBox.Top + k
    MainBox.Left = SelectBox.Right + 12
    MainBox.Top = SelectBox.Top
    MainBox.Right = MainBox.Left + 15
    MainBox.Bottom = SelectBox.Bottom
     
    With cmbPreset
    .AddItem "Custom..."
    .AddItem "Safe Palette (216)"
    End With
    
    'LoadMode pMode
    
    cPaletteIndex = 1

    Me.ForeColor = vbBlack
    If m_WebColors Then
        frmChart.chkWeb.Value = vbChecked
    Else
        frmChart.chkWeb.Value = vbUnchecked
    End If
    
    lblSelColor.BackColor = lblSelColor.BackColor
    LoadLastColor
    Loaded = True
    DefaultDrive = Drive
    DefaultPath = Folder.Path
    IconStart = 0
    Update_Icons
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lColor  As Long
    Dim hDeskDC As Long
    Dim ptMouse As PointAPI

    If mbCapture Then
        If Button = 1 Then
            Call ReleaseCapture
            Call GetCursorPos(ptMouse)
            hDeskDC = GetDC(0)
        
            lColor = GetPixel(hDeskDC, ptMouse.X, ptMouse.Y)
        If lColor <> -1 Then
            mlCurColor = lColor
            Call ResetControls
        End If
        
    Me.MousePointer = vbDefault
    chkCapture.Value = vbUnchecked
        If Me.Visible = False Then
            Me.Visible = True
        End If
    End If
    End If
   
    If Mode = Picker Then
        If X >= SelectBox.Left And X <= SelectBox.Right And Y >= SelectBox.Top And Y <= SelectBox.Bottom Then
            '// In SelectBox Boundary
            SelectBoxHit = True
            Me.MousePointer = vbCustom
            Call MouseOnSelectBox(X, Y)
        End If
        If X >= MainBox.Left And X < MainBox.Right + 11 And Y >= MainBox.Top - 2 And Y < MainBox.Bottom + 3 Then
            '// In MainBox Boundary
            MainBoxHit = True
            If Y > MainBox.Bottom Then Y = MainBox.Bottom
            If Y < MainBox.Top Then Y = MainBox.Top
            Call MouseOnMainBox(X, Y)
        End If
        Else
            If Mode <> About Then
                HandlePresetValues X, Y, Mode
            End If
        End If
    
End Sub

Sub DrawSafePicker(Index As Integer, Clear As Boolean)
    
    Dim r As RECT
    Dim L As Long
    Me.FillStyle = 1
    Me.DrawMode = 13
    Me.DrawWidth = 3
    If Clear Then
        Me.ForeColor = Me.BackColor
        Rectangle Me.hDC, Preset(Index).Left - 2, Preset(Index).Top - 2, Preset(Index).Right + 2, Preset(Index).Bottom + 2
    Else
        r.Left = Preset(Index).Left - 3
        r.Top = Preset(Index).Top - 3
        r.Right = Preset(Index).Right + 3
        r.Bottom = Preset(Index).Bottom + 3
        Call DrawEdge(frmChart.hDC, r, BDR_SUNKENINNER Or BDR_SUNKENOUTER, BF_RECT Or BF_SOFT)
    End If
    
End Sub

Private Sub MouseOnSelectBox(X As Single, Y As Single)

    Dim cl As Long
    Dim hs As HSB
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    DrawPicker
    SelectedPos.X = X
    SelectedPos.Y = Y
    DrawPicker
    cl = GetPixel(Me.hDC, X, Y)
    GetRGB cl, r, g, b
    
    If optS.Value Then
        txtB.Text = Int((255 - SelectedPos.Y + SelectBox.Top) * 100 / 255) 'Brightness Level
        txtH.Text = Int((SelectedPos.X - SelectBox.Left) * 360 / 255) 'Hue Level
        hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
        HSBtoRGB hs, r, g, b
        LoadMainSaturation frmChart.hDC, r, g, b, Val(txtB.Text) / 100
    End If
            
    If optB.Value Then
        hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
        HSBtoRGB hs, r, g, b
        LoadMainBrightness frmChart.hDC, r, g, b
        txtS.Text = Int((SelectBox.Bottom - SelectedPos.Y) * 100 / 255)    ' Saturation Level
        txtH.Text = Int((SelectedPos.X - SelectBox.Left) * 360 / 255)   'Hue Level
    End If

    If optH.Value Then
        txtS.Text = Int((SelectedPos.X - SelectBox.Left) * 100 / 255)   ' Saturation Level
        txtB.Text = Int((255 - SelectedPos.Y + SelectBox.Top) * 100 / 255) 'Brightness Level
    Else
        cl = GetPixel(Me.hDC, MainBox.Left + 5, SelectedMainPos)
        GetRGB cl, r, g, b
    End If
    
    Text1.Text = r
    Text2.Text = g
    Text3.Text = b
    lblSelColor.BackColor = cl
    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    'java
    txtColor(9).Text = "0x" & StrRev(sHex)
    'Hex Color
    sHex = Hex(lColor)
    
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    'CPlus
    txtColor(1).Text = "0x00" & sHex
    'Delphi
    txtColor(8).Text = "$00" & sHex
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
    sGreen = CStr(iGreen)
    sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub MouseOnMainBox(X As Single, Y As Single)
    Dim cl As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer

    DrawSlider SelectedMainPos
    DrawSlider Y
    SelectedMainPos = Y
    cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
    GetRGB cl, r, g, b
            
        If optH.Value Then
            txtH.Text = Int((255 - Y + SelectBox.Top) * 360 / 255)
            GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
        Else
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
            lblSelColor.BackColor = cl
            Dim lColor  As Long
            Dim iIdx    As Integer
            Dim iRed    As Integer
            Dim iGreen  As Integer
            Dim iBlue   As Integer
            Dim sRed    As String
            Dim sGreen  As String
            Dim sBlue   As String
            Dim sHex    As String
            mbNoChange = True
            lColor = lblSelColor.BackColor
            lblSelColor.BackColor = lblSelColor.BackColor
            iRed = (lColor And &HFF&)
            iGreen = (lColor And &HFF00&) / &H100
            iBlue = (lColor And &HFF0000) / &H10000
            txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
            & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
        txtColor(3).Text = CStr(lColor)
    
    'Hex Color
        sHex = Hex(lColor)
        
        If Len(sHex) < 6 Then
            sHex = String(6 - Len(sHex), "0") & sHex
        End If
    'CPlus
    txtColor(1).Text = "0x00" & sHex
    'vbhex
    txtColor(5).Text = "&H" & sHex & "&"
    'Delphi
    txtColor(8).Text = "$00" & sHex
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    'java
    txtColor(9).Text = "0x" & StrRev(sHex)
    sRed = CStr(iRed)
    sGreen = CStr(iGreen)
    sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
            End If
    If optS.Value Then
        txtS.Text = Int((255 - Y + SelectBox.Top) * 100 / 255)
    End If
            
    If optB.Value Then
        txtB.Text = Int((255 - Y + SelectBox.Top) * 100 / 255)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If picOut.Visible = True Then
        Image1.Visible = True
        Image2.Visible = False
     End If
  
    If cmdADD.Visible = True Then
        imgADD.Visible = True
        cmdADD.Visible = False
    End If
    
    imgExit.Visible = True
    cmdExit.Visible = False
    cmdPreset.Visible = False
    imgPreset.Visible = True
    imgCapture.Visible = True
    cmdCapture.Visible = False
    imgHide.Visible = True
    cmdHide.Visible = False
    imageLoad.Visible = True
    cmdLoad.Visible = False
    imgSave.Visible = True
    cmdSave.Visible = False
    imgCopy.Visible = True
    cmdCopy.Visible = False
    imgPrint.Visible = True
    cmdPrint.Visible = False
    imgRect.Visible = True
    cmdRect.Visible = False
    imgHelp.Visible = True
    cmdHelp.Visible = False
    imgPicker.Visible = True
    cmdPicker.Visible = False
    txtColor(1).ForeColor = &H800000
    txtColor(3).ForeColor = &H800000
    txtColor(4).ForeColor = &H800000
    txtColor(5).ForeColor = &H800000
    txtColor(6).ForeColor = &H800000
    txtColor(7).ForeColor = &H800000
    txtColor(8).ForeColor = &H800000
    txtColor(9).ForeColor = &H800000
    lblPantone.ForeColor = &H800000
    
    If Button = 1 Then
        If X >= LcSmpRect.Left And X <= LcSmpRect.Right And Y >= LcSmpRect.Top And Y <= LcSmpRect.Bottom Then
            If pMode < 3 Then
                lblSelColor.BackColor = LastColor
                LoadMode pMode
            End If
            Exit Sub
        End If
    
    If X >= LcButtonRect.Left And X <= LcButtonRect.Right And Y >= LcButtonRect.Top And Y <= LcButtonRect.Bottom Then
        Me.PSet (-100, -100)
        lcButtonPressed = True
        Exit Sub
    End If
        
    If Mode = Picker Then
        If X >= SelectBox.Left And X <= SelectBox.Right And Y >= SelectBox.Top And Y <= SelectBox.Bottom Then
            '// In SelectBox Boundary
            SelectBoxHit = True
            Me.MousePointer = vbCustom
            Call MouseOnSelectBox(X, Y)
        End If
        
        If X >= MainBox.Left And X < MainBox.Right + 11 And Y >= MainBox.Top - 2 And Y < MainBox.Bottom + 3 Then
            '// In MainBox Boundary
            MainBoxHit = True
            If Y > MainBox.Bottom Then Y = MainBox.Bottom
                If Y < MainBox.Top Then Y = MainBox.Top
                    Call MouseOnMainBox(X, Y)
            End If
        Else
            If Mode <> About Then
                HandlePresetValues X, Y, Mode
            End If
        End If
    End If

End Sub

Private Sub HandlePresetValues(X As Single, Y As Single, pMode As sMode)
            
    Dim I As Integer
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    
    For I = 1 To 224
        If X > Preset(I).Left And X < Preset(I).Right And Y > Preset(I).Top And Y < Preset(I).Bottom Then
            DrawSafePicker cPaletteIndex, True
            DrawSafePicker I, False
            Me.PSet (-100, -100)
            cPaletteIndex = I
            
            Select Case pMode
            Case 1
                        
            Case 2
                lblSelColor.BackColor = svdColor(I)
                GetRGB svdColor(I), r, g, b
                GetHexVal r, g, b, HexV
            Case 3
                lblSelColor.BackColor = GetSafeColor(I, r, g, b, HexV)
            End Select
                    
            Exit Sub
        End If
        
        Next I

End Sub

Private Sub PrintRGBHEX(r As Integer, g As Integer, b As Integer, HexV As String)
    
    Me.FontSize = 8
    Me.FillColor = Me.BackColor
    Me.ForeColor = Me.BackColor
    Me.DrawMode = 13
    Me.Line (MainBox.Right + 20, 90)-(MainBox.Right + 20 + 190, 90 + 50), Me.BackColor, BF
    Me.CurrentX = MainBox.Right + 20
    Me.CurrentY = 90
    Me.ForeColor = 0
    Me.Print "R: " & r
    Me.CurrentX = MainBox.Right + 20
    Me.Print "G: " & g
    Me.CurrentX = MainBox.Right + 20
    Me.Print "B: " & b
    Me.CurrentX = MainBox.Right + 20
    Me.Print "Hex: #" & HexV

End Sub

Private Sub GetColorFromHSB(ByVal Sat As Integer, ByVal br As Integer)
    
    '//
    '// This Function Evaluates the Resulting Color Value While Sliding the Hue Shades From Insisting Brightness and Saturation Values
    Dim r As Integer, g As Integer, b As Integer
    Dim Red As Integer, Green As Integer, Blue As Integer
    Dim cl As Long
    Dim X As Long, Y As Long
    On Error Resume Next
    X = Sat * (255 / 100)
    cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
    GetRGB cl, Red, Green, Blue
    r = (Red + ((255 - Red) * (100 - Sat)) / 100) * br / 100
    g = (Green + ((255 - Green) * (100 - Sat)) / 100) * br / 100
    b = (Blue + ((255 - Blue) * (100 - Sat)) / 100) * br / 100
    lblSelColor.BackColor = RGB(r, g, b)

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim em As RECT
    
    em.Right = Screen.Width / 15
    em.Bottom = Screen.Height / 15
    ClipCursor em
    SelectBoxHit = False
    Me.MousePointer = vbDefault
    'PrintLastColor False
    
    If X >= LcSmpRect.Left And X <= LcSmpRect.Right And Y >= LcSmpRect.Top And Y <= LcSmpRect.Bottom Then
         Exit Sub
    End If
    
    If lcButtonPressed And X >= LcButtonRect.Left And X <= LcButtonRect.Right And Y >= LcButtonRect.Top And Y <= LcButtonRect.Bottom Then
         lblSelColor.BackColor = LastColor
         lcButtonPressed = False
    End If
    
    If MainBoxHit And optH.Value Then
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        Dim px As Long
        DrawPicker
        cl = GetPixel(Me.hDC, MainBox.Left + 3, SelectedMainPos)
        GetRGB cl, r, g, b
        LoadVariantsHue r, g, b
        DrawPicker
        cl = Me.Point(SelectedPos.X, SelectedPos.Y)
        px = cl
        GetRGB cl, r, g, b
        Text1.Text = r
        Text2.Text = g
        Text3.Text = b
        lblSelColor.BackColor = Me.Point(SelectedPos.X, SelectedPos.Y)
    End If
    
    If MainBoxHit And optS.Value Then
        DrawPicker
        LoadVariantsSaturation Val(txtS.Text) / 100
        DrawPicker
    End If
    
    If MainBoxHit And optB.Value Then
        DrawPicker
        LoadVariantsBrightness Val(txtB.Text) / 100
        DrawPicker
    End If
    
    Me.PSet (-100, -100)
    MainBoxHit = False
    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    'CPlus
    txtColor(1).Text = "0x00" & sHex
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    'Delphi
    txtColor(8).Text = "$00" & sHex
    'vbhex
    txtColor(5).Text = "&H" & sHex & "&"
    'java
    txtColor(9).Text = "0x" & StrRev(sHex)
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
    sGreen = CStr(iGreen)
    sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    sndPlay "erase", SoundOps.SND_ASYNC
    Dim Last As Integer
    Dim I As Integer
    On Error Resume Next
    Last = FreeFile()
    Open App.Path & "/lastcolor.cps" For Output As #Last
    Print #Last, lblSelColor.BackColor
    Close #Last
    Unload mdi_Help
    Unload frmMagnify
    Unload frmRectangle
    Unload frmPrint
    rs.Close
    Set rs = Nothing
    gConn.Close
    Set gConn = Nothing
    Unload Me
        
End Sub

Private Sub Image2_Click()
    
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If

    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        Picture1.Visible = False
        Picture2.Visible = False
        IconScroll.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If
    
    Me.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    
    If Mode = Picker Then
        Frame2.Visible = False
        Mode = Custom
        chkCapture.Enabled = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        cmbPreset.Visible = True
        LoadCustomColors
        cmbPreset.ListIndex = 1
        lblADDColor.BackColor = lblSelColor.BackColor
        imgADD.Visible = True
        lblADDColor.Visible = True
        Label12.Visible = True
        Label14.Visible = False
        GetRGB lblSelColor.BackColor, r, g, b
        GetHexVal r, g, b, HexV
    Else
        cmbPreset.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        chkCapture.Enabled = True
    
    Select Case True
        Case optH
            optH_Click
        Case optS
            optS_Click
        Case optB
            optB_Click
    End Select
    
    Frame2.Visible = True
    
    If imgADD.Visible = True Then
        imgADD.Visible = False
    End If
    
    lblADDColor.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    Mode = Picker

End If

End Sub

Private Sub Label2_Click()
    
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    lblSelColor.BackColor = Label2.BackColor
    
    If pMode < 3 Then
        LoadMode pMode
    Else
        GetRGB Label2.BackColor, r, g, b
        GetHexVal r, g, b, HexV
        'PrintRGBHEX r, g, b, HexV
    End If

End Sub

Private Sub Label2_DblClick()
    
    
    lblSelColor.BackColor = Label2.BackColor

End Sub

Private Sub lblADDColor_Click()
    
    On Error Resume Next
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblADDColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    
End Sub

Private Sub lblSelColor_Click()
    
    On Error Resume Next
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblSelColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV

End Sub

Private Sub mnuAbout_Click()
    
    frmChart.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim cl As Long
    Dim HexV As String

    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        Picture1.Visible = False
        Picture2.Visible = False
        IconScroll.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
    End If
    
    lblADDColor.Visible = False
    imgADD.Visible = False
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Mode = Custom
    sndPlay "xfilesintro", SoundOps.SND_ASYNC
    Dim iLine As Integer
    picBuffer.Visible = True
    picBackBuffer.Visible = True
    picOut.Visible = True
    Image1.Visible = True
    NumLines = 1
    picBuffer.ScaleMode = vbPixels
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    picBuffer.Visible = False
    Open (App.Path & "\Documents\credits.txt") For Input As #1
    Do Until EOF(1)
        Line Input #1, Tempstring(NumLines)
        NumLines = NumLines + 1
    Loop
    Close #1
    NumLines = NumLines - 1
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    GradiantBackground picBackBuffer
    Me.Picture = LoadPicture(App.Path & "/Images/main3.jpg")
    ReDrawTimer.Interval = 22
    ReDrawTimer.Enabled = True

End Sub

Private Sub mnuColor_Click()

chkCapture.Value = 1

End Sub

Private Sub mnuExit_Click()

    Me.lblSelColor.BackColor = mOldColor
    Unload mdi_Help
    Unload frmMagnify
    Unload frmRectangle
    Unload frmPrint
    Unload Me

End Sub

Private Sub mnuHelp1_Click()

mdi_Help.Show

End Sub

Private Sub mnuRectangle_Click()

    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
    End If

    If CaptureDesktop Then
        '--- Show the Form where all the Work will take place
        DoEvents
        frmRectangle.ShowPicture NewImage.Picture
    End If

    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")

End Sub

Private Sub mnuScreen_Click()

    '--- This is where we capture the FULL screen
    '--- when the User chooses to capture the FULL Screen
    frmChart.Cls
    '--- Capture the Screen
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
    End If
    
    CaptureDesktop
    ScaleBar.Enabled = True
    Mode = Custom
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    IconStart = 0
    Update_Icons
    Picture1.Visible = True
    Picture2.Visible = True
    IconScroll.Visible = True
    lblInfo.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
    Me.Top = 2400
    Me.Show
    
End Sub

Private Sub mnuTrayColor_Click()

chkCapture.Value = 1

End Sub

Private Sub mnuTrayExit_Click()

    sndPlay "erase", SoundOps.SND_ASYNC
    Dim Last As Integer
    Dim I As Integer
    On Error Resume Next
    Last = FreeFile()
    Open App.Path & "/lastcolor.cps" For Output As #Last
    Print #Last, lblSelColor.BackColor
    Close #Last
    Unload mdi_Help
    Unload frmMagnify
    Unload frmRectangle
    Unload frmPrint
    Unload Me

End Sub

Private Sub mnuTrayRectangle_Click()

    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
    End If

    If CaptureDesktop Then
        '--- Show the Form where all the Work will take place
        DoEvents
        frmRectangle.ShowPicture NewImage.Picture
    End If
    
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")

End Sub

Private Sub mnuTrayScreen_Click()
    
    If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
    End If

    frmChart.Cls
    '--- Capture the Screen
    CaptureDesktop
    ScaleBar.Enabled = True
    Mode = Custom
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    IconStart = 0
    Update_Icons
    Picture1.Visible = True
    Picture2.Visible = True
    IconScroll.Visible = True
    lblInfo.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
    Me.Top = 2400
    Me.Show

End Sub

Private Sub mnuTrayShow_Click()
    
    Me.Visible = True
    RemoveFromTray

End Sub

Public Sub optB_Click()
    
    On Error Resume Next
    pMode = 2
    LoadMode pMode

End Sub

Public Sub optH_Click()
    
    On Error Resume Next
    pMode = 0
    LoadMode pMode

End Sub

Public Sub optS_Click()
    
    pMode = 1
    LoadMode pMode

End Sub

Private Sub RedrawTimer_Timer()

    Dim L As Long
    Dim j As Long

    On Error Resume Next
        
    ' Draw the background to the buffer. It's only had to be written once, so we'll just re-blit it over again and agin.
    L = BitBlt(picBuffer.hDC, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hDC, 0, 0, SRCCOPY)
    
    ' Do the following for each line of text in our credits message...
    For j = 1 To NumLines Step 1
        
        ' Set the starting location of where to print the text. Starts off below the bottom of the buffer.
        picBuffer.CurrentY = lY + (j * picBuffer.FontSize + (6 * j))
        picBuffer.CurrentX = (picBuffer.ScaleWidth / 2) - (picBuffer.TextWidth(Tempstring(j)) / 2)
        
        ' Preset the forground color to white
        picBuffer.ForeColor = vbWhite
       
        ' Once the current line of text reaches this point, begin the color shift. This is done for each line
        ' of text in your message
        If picBuffer.CurrentY < 245 Then
            
            ' If a piece of text is color shifting, but not quite to the top yet...
            If picBuffer.CurrentY > 15 Then
                
                ' This changes the forground color to a shade of whatever color (in this case...blue). As
                ' it nears the top, the rate of the R,G anf B values shift differently, to allow a gradual
                ' color shift.
                picBuffer.ForeColor = RGB((((255 / 235) * picBuffer.CurrentY)), (((255 / 235) * picBuffer.CurrentY)), (((255 / 25) * picBuffer.CurrentY)))
            Else
                
                ' We've reached the top...just paint it black and get it over with....
                picBuffer.ForeColor = vbBlack
                
                If j = NumLines And picBuffer.CurrentY < -25 Then
                'If we've painted the last line, and it's above the top, there's no more text to scroll
                ' and we exit.
                    ReDrawTimer.Enabled = False
                    sndPlay "bloop_x", SoundOps.SND_ASYNC
                    'frmChart.Visible = True
                    'Unload Me
                    sndPlay "bloop_x", SoundOps.SND_ASYNC
    picBuffer.Visible = False
    picBackBuffer.Visible = False
    picOut.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    imgPreset.Visible = True
    cmdPreset.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    Me.Cls
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String

    If Mode = Picker Then
        Frame2.Visible = False
        Mode = Custom
        chkCapture.Enabled = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        cmbPreset.Visible = True
        LoadCustomColors
        cmbPreset.ListIndex = 1
        lblADDColor.BackColor = lblSelColor.BackColor
        imgADD.Visible = True
        lblADDColor.Visible = True
        Label12.Visible = True
        Label14.Visible = False
        GetRGB lblSelColor.BackColor, r, g, b
        GetHexVal r, g, b, HexV
    Else
        cmbPreset.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        chkCapture.Enabled = True
    
    Select Case True
    Case optH
         optH_Click
    Case optS
        optS_Click
    Case optB
        optB_Click
    End Select
        
        Frame2.Visible = True
        imgADD.Visible = False
        lblADDColor.Visible = False
        Label12.Visible = False
        Label14.Visible = False
        Mode = Picker
    End If

            End If
        End If
    End If
        
        ' Send the text directly into the buffer hDC
        picBuffer.Print Tempstring(j)
        
    Next
    
    ' Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    ' finished image directly to our output picturebox control.
    L = BitBlt(picOut.hDC, 0, picOut.ScaleTop, picOut.ScaleWidth, picOut.ScaleHeight, picBuffer.hDC, 0, 0, SRCCOPY)
    
    picOut.Refresh
    
    ' Change the offset for the location of where the text will display next turn
    lY = lY - 1

End Sub

Private Sub Text1_Change()
    
    On Error Resume Next
    Dim cm As CMYK
    
    If Text1.Text = "" Then Exit Sub
        cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
        txtCyan.Text = cm.Cyan
        txtMagenta.Text = cm.Magenta
        txtYellow.Text = cm.Yellow
        txtK.Text = cm.k
            If ChangeByInput = False Then Exit Sub
                If Text1.Text <> Int(Text1.Text) Or Val(Text1.Text) < 0 Or Text1.Text > 255 Then
        Exit Sub
    End If
    
    SetColors
    ChangeByInput = False
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub
Private Sub SetColors()
    'On Error Resume Next

    Dim Red As Integer, Green As Integer, Blue As Integer
    Red = Val(Text1.Text)
    Green = Val(Text2.Text)
    Blue = Val(Text3.Text)
    Dim hs As HSB
    hs = RGBtoHSB(RGB(Red, Green, Blue))
    txtH.Text = Int(hs.Hue)
    txtS.Text = Int(hs.Saturation)
    txtB.Text = Int(hs.Brightness)
    lblSelColor.BackColor = RGB(Red, Green, Blue)
    ReloadColors
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

ChangeByInput = True

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    
    If Text1.Text = "" Or Val(Text1.Text) <> Int(Val(Text1.Text)) Or Val(Text1.Text) < 0 Or Val(Text1.Text) > 255 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    Text1.Text = Val(Text1.Text)

End Sub

Private Sub Text2_Change()
    
    If Text2.Text = "" Then Exit Sub
        Dim cm As CMYK
        cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
        txtCyan.Text = cm.Cyan
        txtMagenta.Text = cm.Magenta
        txtYellow.Text = cm.Yellow
        txtK.Text = cm.k
            If ChangeByInput = False Then Exit Sub
                If Text2.Text <> Int(Text2.Text) Or Val(Text2.Text) < 0 Or Text2.Text > 255 Then
            Exit Sub
    End If
    
    SetColors
    ChangeByInput = False
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

ChangeByInput = True

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    
    If Text2.Text = "" Or Val(Text2.Text) <> Int(Val(Text2.Text)) Or Val(Text2.Text) < 0 Or Val(Text2.Text) > 255 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    
    Text2.Text = Val(Text2.Text)

End Sub

Private Sub Text3_Change()
    
    If Text3.Text = "" Then Exit Sub
        Dim cm As CMYK
        If Text3.Text <> Int(Text3.Text) Or Val(Text3.Text) < 0 Or Text3.Text > 255 Then
        Exit Sub
    End If

    cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
    txtCyan.Text = cm.Cyan
    txtMagenta.Text = cm.Magenta
    txtYellow.Text = cm.Yellow
    txtK.Text = cm.k
    If ChangeByInput = False Then Exit Sub
    SetColors
    ChangeByInput = False
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

ChangeByInput = True

End Sub

 Private Sub Text3_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    
    If Text3.Text = "" Or Val(Text3.Text) <> Int(Val(Text3.Text)) Or Val(Text3.Text) < 0 Or Val(Text3.Text) > 255 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    
    Text3.Text = Val(Text3.Text)

End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub txtB_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrB
    AdjustBrightness txtB.Text
    Exit Sub
ErrB:
    sndPlay "BOING", SoundOps.SND_ASYNC
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtB.SetFocus
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Sub DrawPicker()

    Me.DrawMode = 6
    Me.FillStyle = 1
    Me.DrawWidth = 1
    Me.DrawStyle = 0
    Me.Circle (SelectedPos.X, SelectedPos.Y), 5

End Sub

Sub DrawSelFrame()
    
    Dim SelFrame As RECT
    SelFrame.Left = SelectBox.Left - 1
    SelFrame.Top = SelectBox.Top - 1
    SelFrame.Right = SelectBox.Right + 3
    SelFrame.Bottom = SelectBox.Bottom + 3
    DrawEdge Me.hDC, SelFrame, BDR_SUNKENOUTER, BF_RECT
    Me.PSet (-100, -100)

End Sub

Private Sub UpdateColorValues()
        
    Dim r As Integer, g As Integer, b As Integer
    Dim cl As Long
    cl = Me.Point(SelectedPos.X, SelectedPos.Y)
    lblSelColor.BackColor = cl
    GetRGB cl, r, g, b
    Text1.Text = r
    Text2.Text = g
    Text3.Text = b

End Sub

Private Sub txtB_Validate(Cancel As Boolean)
    
    If txtB.Text = "" Or Val(txtB.Text) <> Int(Val(txtB.Text)) Or Val(txtB.Text) < 0 Or Val(txtB.Text) > 100 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 100."
        Cancel = True
        Exit Sub
    End If
    
    txtB.Text = Val(txtB.Text)
    BrightnessEntering = False

End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub txtH_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrH
    AdjustHue txtH.Text
    Exit Sub
ErrH:
    sndPlay "BOING", SoundOps.SND_ASYNC
    MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
    txtH.SetFocus

End Sub

Private Sub txtH_Validate(Cancel As Boolean)
    
    If txtH.Text = "" Or Val(txtH.Text) <> Int(Val(txtH.Text)) Or Val(txtH.Text) < 0 Or Val(txtH.Text) > 360 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 360."
        Cancel = True
        Exit Sub
    End If
    
    txtH.Text = Val(txtH.Text)
  
End Sub

Private Sub txtK_Change()
    
    If ChangeByInput = False Then Exit Sub
    If txtK.Text <> Int(txtK.Text) Or Val(txtK.Text) < 0 Or txtK.Text > 255 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 255."
        Exit Sub
    End If

End Sub

Private Sub txtK_KeyPress(KeyAscii As Integer)

ChangeByInput = True

End Sub

Private Sub txtK_Validate(Cancel As Boolean)
    
    If txtK.Text <> Int(txtK.Text) Or Val(txtK.Text) < 0 Or txtK.Text > 255 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 255."
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub txtS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub txtS_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrS
    AdjustSaturation txtS.Text
    Exit Sub
ErrS:
    sndPlay "BOING", SoundOps.SND_ASYNC
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtS.SetFocus
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub AdjustHue(ByVal Hue As Single)
        
    Dim cl As Long
    Dim hs As HSB
    Dim r As Integer, g As Integer, b As Integer
    
    If Hue > 360 Or Hue < 0 Or Int(Hue) <> Hue Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
        txtH.SetFocus
        txtH.Text = OldHue
        AdjustHue OldHue
        Exit Sub
    End If
    
    DrawPicker
    DrawSlider SelectedMainPos
    Select Case True
    
        Case optH.Value
            SelectedMainPos = MainBox.Bottom - Hue * (255 / 360)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, b
            LoadVariantsHue r, g, b
            UpdateColorValues
        Case optS.Value
            SelectedPos.X = 10 + Hue * (255 / 360)
            hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
            HSBtoRGB hs, r, g, b
            LoadMainSaturation frmChart.hDC, r, g, b, Val(txtB.Text) / 100
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        Case optB.Value
            SelectedPos.X = SelectBox.Left + Hue * (255 / 360)
            hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            HSBtoRGB hs, r, g, b
            LoadMainBrightness frmChart.hDC, r, g, b
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
    
    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
        
End Sub

Private Sub AdjustSaturation(ByVal Saturation As Single)
        
    Dim cl As Long
    Dim r As Integer, g As Integer, b As Integer
    Dim hs As HSB
    If Saturation > 100 Or Saturation < 0 Or Int(Saturation) <> Saturation Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
        txtS.SetFocus
        txtS.Text = OldSaturation
        AdjustSaturation OldSaturation
        Exit Sub
    End If

    DrawPicker
    DrawSlider SelectedMainPos
    Select Case True
        Case optH.Value
            SelectedPos.X = SelectBox.Left + Saturation * (255 / 100)
            UpdateColorValues
        Case optS.Value
            LoadVariantsSaturation Val(txtS.Text) / 100
            SelectedMainPos = MainBox.Bottom - Saturation * (255 / 100)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        Case optB.Value
            SelectedPos.Y = SelectBox.Bottom - Saturation * (255 / 100)
            hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            GetRGB cl, r, g, b
            HSBtoRGB hs, r, g, b
            LoadMainBrightness frmChart.hDC, r, g, b
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
        Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
        
End Sub

Private Sub AdjustBrightness(ByVal Brightness As Single)
    
    Dim r As Integer, g As Integer, b As Integer
    Dim cl As Long
    Dim hs As HSB
    If Brightness > 100 Or Brightness < 0 Or Int(Brightness) <> Brightness Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
        txtB.Text = OldBrightness
        txtB.SetFocus
        AdjustBrightness OldBrightness
        Exit Sub
    End If

    DrawPicker
    DrawSlider SelectedMainPos

    Select Case True
        Case optH.Value
            SelectedPos.Y = SelectBox.Bottom - Brightness * (255 / 100)
            UpdateColorValues
        Case optS.Value
            SelectedPos.Y = SelectBox.Bottom - Brightness * (255 / 100)
            hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
            HSBtoRGB hs, r, g, b
            LoadMainSaturation frmChart.hDC, r, g, b, Val(txtB.Text) / 100
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
            
        Case optB.Value
            SelectedMainPos = MainBox.Bottom - Brightness * (255 / 100)
            LoadVariantsBrightness Brightness / 100
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            GetRGB cl, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
    
    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String
    mbNoChange = True
    lColor = lblSelColor.BackColor
    lblSelColor.BackColor = lblSelColor.BackColor
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    mbNoChange = False
        
End Sub

Private Sub txtS_Validate(Cancel As Boolean)
    
    If txtS.Text = "" Or Val(txtS.Text) <> Int(Val(txtS.Text)) Or Val(txtS.Text) < 0 Or Val(txtS.Text) > 100 Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "Enter an Integer between 0 and 100."
        Cancel = True
        Exit Sub
    End If
    
    txtS.Text = Val(txtS.Text)
    
End Sub

Private Function ConvertToSysColor(ByVal lColor As Long) As Long

'Find a system color that matches lColor

    Dim lIdx As Long
    Dim sHex As String

    If lColor < 0 Then
        'Already a system color
        ConvertToSysColor = lColor
    Else
        For lIdx = 0 To 24
            If GetSysColor(lIdx) = lColor Then
                'Found a match
                sHex = Hex$(lIdx)
                If Len(sHex) < 2 Then
                    sHex = "0" & sHex
                End If
                ConvertToSysColor = Val("&H800000" & sHex)
                Exit For
            End If
        Next
        
        If lIdx > 24 Then
            'Didn't find a match
            ConvertToSysColor = -1
        End If
    End If
    
End Function

Private Sub DrawButtonUp(iIdx As Integer)

    'Draw the button in the "unpressed" position

    Dim iX          As Integer
    Dim iY          As Integer
    Dim iBtnWidth   As Integer
    Dim iBtnHeight  As Integer

    If iIdx >= 0 Then
        iBtnWidth = picQBColors.ScaleWidth / 2
        iBtnHeight = picQBColors.ScaleHeight / 8
        iX = Int(iIdx / 8) * iBtnWidth
        iY = (iIdx Mod 8) * iBtnHeight
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vbButtonFace, BF
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vb3DDKShadow, B
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 2, iBtnHeight - 2), vb3DHighlight, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 3, iBtnHeight - 3), vbButtonShadow, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 4, iBtnHeight - 4), vbButtonFace, BF
        picQBColors.Line (iX + 4, iY + 4)-Step(iBtnWidth - 10, iBtnHeight - 10), QBColor(iIdx), BF
        picQBColors.Line (iX + 4, iY + 4)-Step(iBtnWidth - 10, iBtnHeight - 10), &H0&, B
        picQBColors.CurrentX = Int(((iX + (iX + iBtnWidth - 1)) / 2) - (picQBColors.TextWidth(CStr(iIdx)) / 2))
        picQBColors.CurrentY = Int(((iY + (iY + iBtnHeight - 1)) / 2) - (picQBColors.TextHeight(CStr(iIdx)) / 2))
        picQBColors.ForeColor = Abs(CInt(iIdx < 10)) * &HFFFFFF
        picQBColors.Print CStr(iIdx)
        miCurQBIdx = -1
    End If
    
End Sub

Private Sub DrawButtonDown(iIdx As Integer)

'Draw the button in the "pressed" position

    Dim iX          As Integer
    Dim iY          As Integer
    Dim iBtnWidth   As Integer
    Dim iBtnHeight  As Integer
    Dim sConst      As String

    If iIdx >= 0 Then
        iBtnWidth = picQBColors.ScaleWidth / 2
        iBtnHeight = picQBColors.ScaleHeight / 8
        iX = Int(iIdx / 8) * iBtnWidth
        iY = (iIdx Mod 8) * iBtnHeight
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vbButtonFace, BF
        picQBColors.Line (iX, iY)-Step(iBtnWidth - 1, iBtnHeight - 1), vb3DDKShadow, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 3, iBtnHeight - 3), vb3DHighlight, B
        picQBColors.Line (iX + 1, iY + 1)-Step(iBtnWidth - 4, iBtnHeight - 4), vbButtonShadow, B
        picQBColors.Line (iX + 2, iY + 2)-Step(iBtnWidth - 5, iBtnHeight - 5), vbButtonFace, B
        picQBColors.Line (iX + 5, iY + 5)-Step(iBtnWidth - 10, iBtnHeight - 10), QBColor(iIdx), BF
        picQBColors.Line (iX + 5, iY + 5)-Step(iBtnWidth - 10, iBtnHeight - 10), &H0&, B
        picQBColors.CurrentX = Int(((iX + (iX + iBtnWidth - 1)) / 2) - (picQBColors.TextWidth(CStr(iIdx)) / 2)) + 1
        picQBColors.CurrentY = Int(((iY + (iY + iBtnHeight - 1)) / 2) - (picQBColors.TextHeight(CStr(iIdx)) / 2)) + 1
        picQBColors.ForeColor = Abs(CInt(iIdx < 10)) * &HFFFFFF
        picQBColors.Print CStr(iIdx)
        miCurQBIdx = iIdx
        mlCurColor = QBColor(iIdx)
    End If
    
End Sub

Private Sub DrawQBButtons()

'Draw all the QBColor buttons

    Dim iIdx As Integer

    For iIdx = 0 To 15
        miCurQBIdx = 0
        Call DrawButtonUp(iIdx)
    Next iIdx

End Sub

Private Sub GetSystemColors()

    Dim iIdx    As Integer
    Dim aColors As Variant

    'Setup the Text and VB constants for the system colors.
    aColors = Array(Array("Scroll Bars", "Desktop", "Active Title Bar", "Inactive Title Bar", "Menu Bar", "Window Background", _
        "Window Frame", "Menu Text", "Window Text", "Active Title Bar Text", "Active Border", "Inactive Border", _
        "Application Workspace", "Highlight", "Highlight Text", "Button Face", "Button Shadow", "Disabled Text", "Button Text", _
        "Inactive Title Bar Text", "Button Highlight", "Button Dark Shadow", "Button Light Shadow", "ToolTip Text", "ToolTip Background"), _
        Array("vbScrollBars", "vbDesktop", "vbActiveTitleBar", "vbInactiveTitleBar", "vbMenuBar", "vbWindowBackground", _
        "vbWindowFrame", "vbMenuText", "vbWindowText", "vbActiveTitleBarText", "vbActiveBorder", "vbInactiveBorder", _
        "vbApplicationWorkspace", "vbHighlight", "vbHighlightText", "vbButtonFace", "vbButtonShadow", "vbGrayText", "vbButtonText", _
        "vbInactiveTitleBarText", "vb3DHighlight", "vb3DDKShadow", "vb3DLight", "vbInfoText", "vbInfoBackground"))
        
    ReDim msaSysColors(UBound(aColors(0)))
    For iIdx = 0 To UBound(aColors(0))
        cboSysColors.AddItem aColors(0)(iIdx)
        msaSysColors(iIdx) = aColors(1)(iIdx)
    Next
    
    Erase aColors
    
End Sub

Private Sub ResetControls()

'Set all controls to their correct values
'based on the current color setting (mlCurColor).

    Dim lColor  As Long
    Dim iIdx    As Integer
    Dim iRed    As Integer
    Dim iGreen  As Integer
    Dim iBlue   As Integer
    Dim sRed    As String
    Dim sGreen  As String
    Dim sBlue   As String
    Dim sHex    As String

    'Stop controls from triggering events that would call this
    'procedure again while changing control properties here.
    mbNoChange = True
    
    'Find out if it's a system color or it matches
    'a system color (Returns -1 if it's not).
    lColor = ConvertToSysColor(mlCurColor)
    
    'Convert system color, if needed.
    'A Long of &H80000000& or greater is a negative number.
    'System colors in VB are &H80000000& to &H80000018& (-2147483648 to -2147483624).
    'All other colors are &H00000000& to &H00FFFFFF& (0 to 16777215).
    If lColor < -1 Then  'If it's a System color...
        cboSysColors.ListIndex = (lColor And &HFF&)
        txtColor(6).Text = msaSysColors(cboSysColors.ListIndex)
        txtColor(7).Text = "&H" & Hex$(lColor) & "&"
        '(&H8000000F& And &HFF&) = &HF(15), which is the real system
        'color index for ButtonFace. GetSysColor() will return the
        'actual color setting for the system based on this index.
        lColor = GetSysColor(lColor And &HFF&)
    Else
        'Not a system color
        lColor = mlCurColor
        txtColor(6).Text = "N/A"
        txtColor(7).Text = "N/A"
        cboSysColors.ListIndex = -1
    End If
    
    'Show the color in the Color box.
    lblSelColor.BackColor = lColor
    
    'VB QBColor and Constants
    Select Case lColor
        Case vbBlack
            txtColor(0).Text = "vbBlack"
            
        Case vbBlue
            txtColor(0).Text = "vbBlue"
            
        Case vbGreen
            txtColor(0).Text = "vbGreen"
            
        Case vbCyan
            txtColor(0).Text = "vbCyan"
            
        Case vbRed
            txtColor(0).Text = "vbRed"
            
        Case vbMagenta
            txtColor(0).Text = "vbMagenta"
            
        Case vbYellow
            txtColor(0).Text = "vbYellow"
            
        Case vbWhite
            txtColor(0).Text = "vbWhite"
            
        Case Else
            txtColor(0).Text = "N/A"
            
    End Select
    
    'RGB Color - extract Red, Green and Blue values.
    iRed = (lColor And &HFF&)
    iGreen = (lColor And &HFF00&) / &H100
    iBlue = (lColor And &HFF0000) / &H10000
    txtColor(2).Text = "RGB(" & CStr(iRed) & ", " _
        & CStr(iGreen) & ", " & CStr(iBlue) & ")"
    
    'Decimal Color
    txtColor(3).Text = CStr(lColor)
    
    'Hex Color
    sHex = Hex(lColor)
    If Len(sHex) < 6 Then
        sHex = String(6 - Len(sHex), "0") & sHex
    End If
    txtColor(5).Text = "&H" & sHex & "&"
    
    'HTML Color - Reverse Hex
    txtColor(4).Text = "#" & Right$(sHex, 2) & Mid$(sHex, 3, 2) & Left$(sHex, 2)
    
    'Enable/disable the TextBoxes based on whether
    'or not they contain "N/A".
    For iIdx = 0 To txtColor.UBound
        '(InStr(1, txtColor(iIdx).Text, "N/A") = 0) = True if "N/A" is NOT present
        txtColor(iIdx).Enabled = (InStr(1, txtColor(iIdx).Text, "N/A") = 0)
    Next
    
    'Set the QBColor buttons
    'If InStr(1, txtColor(1).Text, "N/A") = 0 Then
        'iIdx = Val(Mid$(txtColor(1).Text, 9))
        If iIdx <> miCurQBIdx Then
            Call DrawButtonUp(miCurQBIdx)
            Call DrawButtonDown(iIdx)
        End If
    'Else
        'Call DrawButtonUp(miCurQBIdx)
    'End If
    
    'Setup the color values for the RGB labels
    If chkHex.Value = vbChecked Then
        sRed = Hex(iRed)
        sGreen = Hex(iGreen)
        sBlue = Hex(iBlue)
    Else
        sRed = CStr(iRed)
        sGreen = CStr(iGreen)
        sBlue = CStr(iBlue)
    End If
    
    'Set the RGB and Shade scroll bars
    hsbRGB(0).Value = iRed
    hsbRGB(1).Value = iGreen
    hsbRGB(2).Value = iBlue
    hsbShade.Value = (iRed + iGreen + iBlue) / 3
    
    'Set the RGB TextBoxes if they changed
    If txtRGB(0).Text <> sRed Then
        txtRGB(0).Text = sRed
        txtRGB(0).SelStart = Len(txtRGB(0).Text)
    End If
    If txtRGB(1).Text <> sGreen Then
        txtRGB(1).Text = sGreen
        txtRGB(1).SelStart = Len(txtRGB(1).Text)
    End If
    If txtRGB(2).Text <> sBlue Then
        txtRGB(2).Text = sBlue
        txtRGB(2).SelStart = Len(txtRGB(2).Text)
    End If
    
    'Allow control events again
    lblSelColor.BackColor = lblSelColor.BackColor
    mbNoChange = False
    Dim Red As Integer, Green As Integer, Blue As Integer
    Red = Val(txtRGB(0).Text)
    Green = Val(txtRGB(1).Text)
    Blue = Val(txtRGB(2).Text)
    Text1.Text = Val(txtRGB(0).Text)
    Text2.Text = Val(txtRGB(1).Text)
    Text3.Text = Val(txtRGB(2).Text)
    Dim hs As HSB
    hs = RGBtoHSB(RGB(Red, Green, Blue))
    'Call DrawSlider(SelectedMainPos)
    txtH.Text = Int(hs.Hue)
    txtS.Text = Int(hs.Saturation)
    txtB.Text = Int(hs.Brightness)
    lblSelColor.BackColor = RGB(Red, Green, Blue)
    ReloadColors
    If frmMagnify.Visible = True Then
        frmMagnify.Visible = False
    End If
    
End Sub

Private Sub cboSysColors_Click()
If picOut.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        picBuffer.Visible = False
        picBackBuffer.Visible = False
        picOut.Visible = False
        Image1.Visible = False
        Image2.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If

    If Picture2.Visible = True Then
        sndPlay "bloop_x", SoundOps.SND_ASYNC
        Picture1.Visible = False
        Picture2.Visible = False
        IconScroll.Visible = False
        lblInfo.Visible = False
        fraQBColors.Visible = True
        frmChart.Label8.Visible = True
        VMove.Visible = False
        HMove.Visible = False
        Label13.Visible = False
        ScaleBar.Visible = False
        ScaleMeter.Visible = False
        imgPreset.Visible = True
        cmdPreset.Visible = False
        Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
    End If
    
    Me.Cls
    'Dim r As Integer, g As Integer, b As Integer
    'Dim HexV As String
    
    'If Mode = Picker Then
        'Frame2.Visible = False
        'Me.Picture = LoadPicture(App.Path & "/Images/main2.jpg")
        'Mode = Custom
        'chkCapture.Enabled = False
        'Label12.Visible = True
        'Label14.Visible = False
        'imgPreset.Visible = True
        'cmdPreset.Visible = False
        'cmbPreset.Visible = True
        'LoadCustomColors
        'cmbPreset.ListIndex = 1
        'lblADDColor.BackColor = lblSelColor.BackColor
        'GetRGB lblSelColor.BackColor, r, g, b
        'GetHexVal r, g, b, HexV

    'Else
        'cmbPreset.Visible = False
        'chkCapture.Enabled = True
        'Me.Picture = LoadPicture(App.Path & "/Images/main1.jpg")
        
    'Select Case True
        'Case optH
            'optH_Click
        'Case optS
            'optS_Click
        'Case optB
            'optB_Click
    'End Select
    
    'Frame2.Visible = True
    'imgADD.Visible = False
    'lblADDColor.Visible = False
    'Label12.Visible = False
    'Label14.Visible = False
    'imgPreset.Visible = True
    'cmdPreset.Visible = False
    Mode = Picker
    
    'End If
    Dim sHex    As String

    'If change not coming from ResetControls
    If Not mbNoChange Then
        If cboSysColors.ListIndex >= 0 Then
            sHex = Hex(cboSysColors.ListIndex)
            If Len(sHex) < 2 Then
                sHex = "0" & sHex
            End If
            mlCurColor = Val("&H800000" & sHex)
        End If
        Call ResetControls
    End If
    
End Sub

Private Sub chkHex_Click()

    Dim iIdx As Integer
    Dim iMax As Integer

    If chkHex.Value = vbChecked Then
        iMax = 2
    Else
        iMax = 3
    End If
    
    For iIdx = 0 To 2
        txtRGB(iIdx).MaxLength = iMax
    Next
    
    Call ResetControls

End Sub

Private Sub chkMsg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim wParam As Long
    
    'X contains the wParam from the Windows messaging system
    'when a mouse event occurs over the System Tray Icon.
    'Since the checkbox coordinates are in twips, VB has
    'already converted the message by multiplying the X by
    'Screen.TwipsPerPixelX, so convert it back.
    wParam = X / Screen.TwipsPerPixelX
    
    'This is only using Double-Click and Right-Click,
    'but all of the following events are returned.
    Select Case wParam
        Case WM_MOUSEMOVE
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            If Forms.count > 1 Then
                Call AppActivate(Me.Caption)
            Else
                Me.Show
            End If
        Case WM_RBUTTONDOWN
            'ShowPopupAtCursor Me, mnuFile, vbPopupMenuRightButton + vbPopupMenuRightAlign, mnuFile_Show
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select

End Sub


Private Sub chkCapture_Click()

    If chkCapture.Value = 1 Then
        frmMagnify.Visible = True
        Timer1.Interval = 10
        Timer1.Enabled = True
        mbCapture = True
        Me.MousePointer = vbCrosshair
        Call ReleaseCapture
        Call SetCapture(Me.hwnd)
    End If
    
End Sub

Private Sub cmdHide_Click()
    
    Me.Hide
    Mod_Tray.AddToTray Me, mnuTray
    Mod_Tray.AddToTrayToolTip Me, mnuTray, "Right click me for a menu", "ColorWiz", 2

End Sub

Private Sub Form_Paint()

    Dim iIdx    As Integer
    Dim iSavIdx As Integer
    Dim lLeft   As Long
    Dim lTop    As Long
    Dim lRight  As Long
    Static bBusy As Boolean

    'Don't allow re-entry caused by our own painting
    If Not bBusy Then
        bBusy = True
        lLeft = fraQBColors.Left - Screen.TwipsPerPixelX
        lRight = txtColor(3).Left + txtColor(3).Width + (2 * Screen.TwipsPerPixelX)
        
        lTop = lblCap(0).Top - ((lblCap(0).Top - (fraQBColors.Top + fraQBColors.Height)) / 2)
        lTop = Int(lTop / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
        Me.Line (lLeft, lTop)-(lRight, lTop), vb3DShadow
        Me.Line (lLeft, lTop + Screen.TwipsPerPixelY)-(lRight, lTop + Screen.TwipsPerPixelY), vb3DHighlight
        
'        lTop = Me.ScaleHeight - ((Me.ScaleHeight - (txtColor(4).Top + txtColor(4).Height)) / 2)
'        lTop = Int(lTop / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
'        Me.Line (lLeft, lTop)-(lRight, lTop), vb3DShadow
'        Me.Line (lLeft, lTop + Screen.TwipsPerPixelY)-(lRight, lTop + Screen.TwipsPerPixelY), vb3DHighlight
        
        iSavIdx = miCurQBIdx
        For iIdx = 0 To 15
            Call DrawButtonUp(iIdx)
        Next iIdx
        miCurQBIdx = iSavIdx
        If miCurQBIdx >= 0 Then
            Call DrawButtonDown(miCurQBIdx)
        End If
        
        bBusy = False
    End If

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        Me.Hide
        Me.WindowState = vbNormal
    End If
    
End Sub

Private Sub hsbRGB_Change(Index As Integer)

    If Mode = SafeColor Then
        Exit Sub
    End If
    'If change not coming from ResetControls
    If Not mbNoChange Then
        mlCurColor = RGB(hsbRGB(0).Value, hsbRGB(1).Value, hsbRGB(2).Value)
        Call ResetControls
    End If
    
End Sub

Private Sub hsbRGB_Scroll(Index As Integer)
    
    Call hsbRGB_Change(Index)

End Sub

Private Sub hsbShade_Change()

    Dim iRed        As Integer
    Dim iGreen      As Integer
    Dim iBlue       As Integer
    Dim iChange     As Integer
    Dim lColor      As Long
    Static iOldVal  As Integer

    'If change not coming from ResetControls
    If Not mbNoChange Then
        iChange = hsbShade.Value - iOldVal
        lColor = RGB(hsbRGB(0).Value, hsbRGB(1).Value, hsbRGB(2).Value)
        iRed = (lColor And &HFF&)
        iGreen = (lColor And &HFF00&) / &H100
        iBlue = (lColor And &HFF0000) / &H10000
        iRed = iRed + iChange
        iGreen = iGreen + iChange
        iBlue = iBlue + iChange
        If iRed > 255 Then
            iRed = 255
        ElseIf iRed < 0 Then
            iRed = 0
        End If
        If iGreen > 255 Then
            iGreen = 255
        ElseIf iGreen < 0 Then
            iGreen = 0
        End If
        If iBlue > 255 Then
            iBlue = 255
        ElseIf iBlue < 0 Then
            iBlue = 0
        End If
        mlCurColor = RGB(iRed, iGreen, iBlue)
        Call ResetControls
    End If

    iOldVal = hsbShade.Value
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub hsbShade_Scroll()
    
    Call hsbShade_Change
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub lblSelColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'frmColor.Show
    
End Sub


Private Sub mnuFile_Exit_Click()
    
    On Error Resume Next
    Unload Me

End Sub

Private Sub mnuFile_Show_Click()
    
    If Forms.count > 1 Then
        Call AppActivate(Me.Caption)
    Else
        Me.Show
    End If

End Sub
Private Sub Form_Activate()

    Timer1.Interval = 10
    Timer1.Enabled = True
    NewImage.Left = 0
    NewImage.Top = 0
    
End Sub

Private Sub picQBColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim iIdx    As Integer
Dim sColor  As String

    If Button = 1 Then
        iIdx = (Int(X / (picQBColors.ScaleWidth / 2)) * 8) + _
            Int(Y / (picQBColors.ScaleHeight / 8))
        If mlCurColor <> QBColor(iIdx) Then
            mlCurColor = QBColor(iIdx)
            Call ResetControls
        End If
    End If
    lblPantone.Caption = "Select a Pantone Color"
    cboPantone.Text = "Pantone Colors"

End Sub

Private Sub txtColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    With txtColor(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtRGB_Change(Index As Integer)

    Dim iIdx    As Integer
    Dim iVal    As Integer
    Dim sChar   As String
    Dim sText   As String
    Static bBusy As Boolean
    Static sOldText(2) As String

    If Not (bBusy Or mbNoChange) Then
        bBusy = True
        With txtRGB(Index)
            For iIdx = 1 To Len(.Text)
                sChar = UCase$(Mid$(.Text, iIdx, 1))
                If sChar >= "0" And sChar <= "9" Then
                    sText = sText & sChar
                ElseIf chkHex.Value = vbChecked Then
                    If sChar >= "A" And sChar <= "F" Then
                        sText = sText & sChar
                    End If
                End If
            Next
            If chkHex.Value = vbChecked Then
                iVal = Val("&H" & sText)
            Else
                iVal = Val(sText)
            End If
            If Len(sText) <> Len(.Text) Or iVal > 255 Then
                Beep
                .Text = sOldText(Index)
                .SelStart = Len(.Text)
            End If
            sOldText(Index) = .Text
            hsbRGB(Index).Value = iVal
        End With
        bBusy = False
    End If
    
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)

    With txtRGB(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txtRGB_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    
    If KeyAscii > 8 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        If chkHex.Value <> vbChecked Then
            Beep
            KeyAscii = 0
        ElseIf KeyAscii < vbKeyA Or KeyAscii > vbKeyF Then
            Beep
            KeyAscii = 0
        End If
    End If
        
End Sub

Private Sub Timer1_Timer()
    
    lblSpyColor.BackColor = spy_color
End Sub

Public Function spy_color() As Long

    On Error Resume Next
    Dim CursorPos As PointAPI
    Dim rDC As Long, rPixel As Long
    Call GetCursorPos(CursorPos)
    rDC& = GetDC(0&)
    rPixel& = GetPixel(rDC&, CursorPos.X, CursorPos.Y)
    Call ReleaseDC(0&, rDC&)
    spy_color& = rPixel&

End Function

Private Function GradiantBackground(picBox As PictureBox)
    
    ipicWidth = picBox.ScaleWidth
    ipicHeight = picBox.ScaleHeight
    iColorCur = 255
    iColorStep = 5 * (0 - 255) / ipicHeight

    For lYOffset = 0 To ipicHeight Step 5
        picBox.Line (-1, lYOffset - 1)-(ipicWidth, lYOffset + 5), RGB(0, 0, iColorCur), BF
        iColorCur = iColorCur + iColorStep
    Next lYOffset
    
End Function

Public Sub LoadPantoneCbo()
    '************************************
    'This sub loads Pantone names in combo
    '************************************
    Dim rs1 As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    'establish recordset connection
    sql = "Select PANTONENAM from tblPantoneConvTable"
    rs1.Open sql, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText

    'populate combo box
    Do Until rs1.EOF
        cboPantone.AddItem rs1("Pantonenam")
        rs1.MoveNext
    Loop
    
    'release rs
    rs1.Close
    Set rs1 = Nothing

Exit Sub
ErrorHandler:
    MsgBox "Error loading Pantone combo box!" & vbNewLine & Err.Number & " - " & Err.Description
    sndPlay "BOING", SoundOps.SND_ASYNC
    Exit Sub
    Resume

End Sub
Private Function CaptureDesktop() As Boolean

'--- 07/30/2004 GCC - Better way to capture screen...(doesn't use Windows clipboard!):
    Static C%
    Dim hWndScreen As Long
    Dim nXpos As Long
    Dim I As Integer
    Dim FileSelect$
    C% = C% + 1
    j% = C%
    '--- Hide the frmChart so that it will not
    '--- be included in the Screen Capture
    '--- NOTE: It seems sometimes that when Windows gets bogged down,
    '--- it "captures" a ghost of frmChart because it didn't fully
    '--- hide it.  Here are some attempts to avoid that:
    '--- To make sure this form is not included in screencapture, let's
    '--- hide it AND move it off to the left of the screen, then move it
    '--- back and show it.
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
        Set NewImage.Picture = CaptureWindow(hWndScreen, False, 0, 0, _
                Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
    End With
    
    Me.Left = nXpos

    CaptureDesktop = True
    FileSelect$ = App.Path & "\Capture\TempImg" & C% & ".bmp"
    SavePicture NewImage.Picture, FileSelect$
    Filename$ = FileSelect$
    Picture_Shown = True
    NewImage.Visible = True
    Sample.Picture = LoadPicture(Filename$)
    NewImage.Picture = LoadPicture(Filename$)
    NewImage_Width = NewImage.Width
    NewImage_Height = NewImage.Height
    ScaleBar.Enabled = True
    ScaleBar.Value = 100
    ScaleBar.Max = 500
    If NewImage_Width * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Width) * 100
    If NewImage_Height * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Height) * 100
    NewImage.Cls
    Picture_Update

End Function
Private Sub ScaleBar_Change()
    
    Picture_Update

End Sub
Private Sub ScaleBar_Scroll()
    
    Picture_Update

End Sub
Private Sub Picture_Update()

    On Error Resume Next
    ScaleMeter.Caption = Mid$(Str$(ScaleBar.Value), 2) + "%"
    WWW = NewImage_Width * (ScaleBar.Value / 100)
    WWS = Sample_Width * (ScaleBar.Value / 100)
        If WWW > 4125 Then NewImage.Width = 4125 Else NewImage.Width = WWW
        HHH = NewImage_Height * (ScaleBar.Value / 100)
            If HHH > 3490 Then NewImage.Height = 3490 Else NewImage.Height = HHH
            VMove.Max = HHH - NewImage.Height
            HMove.Max = WWW - NewImage.Width

    NewImage.PaintPicture Sample, 0, 0, WWW, HHH, 0, 0, NewImage_Width, NewImage_Height, vbSrcCopy

End Sub
Private Sub VMove_Change()

MovePicture

End Sub
Private Sub VMove_Scroll()

MovePicture

End Sub
Private Sub HMove_Change()

If Dragnow = False Then MovePicture

End Sub
Private Sub HMove_Scroll()

MovePicture

End Sub
Private Sub MovePicture()

    S = (ScaleBar.Value / 100)
    NewImage.PaintPicture Sample, 0, 0, NewImage.Width, NewImage.Height, HMove.Value / S, VMove.Value / S, NewImage.Width / S, NewImage.Height / S, vbSrcCopy

End Sub
Private Sub NewImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- This where we set the Begainning of the Box
'--- that will be Drawn around the Capture Area
If ScaleBar.Enabled = False Then Exit Sub
Dragnow = True
DragX = X + HMove.Value
DragY = Y + VMove.Value
    If mbCrop Then
        mbDown = (Button = 1)
        NewImage.MousePointer = vbCrosshair
        
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


Private Sub NewImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--- Where we Draw the Box around the Choosen Area as you hold down the Left Mouse
'--- button and Drag in any direction to create a rectangle
If Dragnow = False Then Exit Sub
10 H = Int(DragX - X)
v = Int(DragY - Y)
If H < 0 Then H = 0 Else If H > HMove.Max Then H = HMove.Max
If v < 0 Then v = 0 Else If v > VMove.Max Then v = VMove.Max
HMove.Value = H
VMove.Value = v
MovePicture
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
Private Sub NewImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    Dim XUpperLeft As Long
    Dim YUpperLeft As Long
    Dim XLowerRight As Long
    Dim YLowerRight As Long
    Dragnow = False
    '--- 07/30/2004 GCC - Only process this MouseUp event during an actual "crop":
    If mbCrop Then
    
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
        NewImage.MousePointer = vbDefault
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
            .PaintPicture NewImage.Picture, 0, 0, _
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
        With NewImage
            '--- Incase picture was scrolled over (via scrollbars), reset it's position
            .Left = 0
            .Top = 0
            
            '--- Just to be safe, clear picture before
            '--- loading new image:
            .Picture = LoadPicture()
            .Cls
            .Picture = picRectangle.Image
        End With
        
    End If
End Sub

Private Sub IconScroll_Change()
IconStart = IconScroll.Value: Update_Icons
End Sub
Private Sub IconScroll_Scroll()
IconStart = IconScroll.Value: Update_Icons
End Sub
Private Sub Update_Icons()
Dim Ratio As Double
For I = 0 To 7
picIcon(I).Visible = False
Label1(I).Visible = False
Next I
b = 0
10 b = b + 1
If b > FileList.ListCount Then Exit Sub
If b > 8 Then Exit Sub
If Right$(Folder.Path, 1) = "\" Then r$ = "" Else r$ = "\"

picIcon(b - 1).Picture = LoadPicture(Folder.Path + r$ + FileList.List(b - 1 + IconStart))
Ratio = picIcon(b - 1).Picture.Width / picIcon(b - 1).Picture.Height
If picIcon(b - 1).Picture.Height >= picIcon(b - 1).Picture.Width Then
            picIcon(b - 1).Height = Pic2(0).Height
            picIcon(b - 1).Width = Pic2(0).Height * Ratio
        Else
            picIcon(b - 1).Width = Pic2(0).Width
            picIcon(b - 1).Height = Pic2(0).Width / Ratio
        End If
    Label1(0).Top = picIcon(0).Top + picIcon(0).Height + 20
    Label1(1).Top = picIcon(1).Top + picIcon(1).Height + 20
    Label1(2).Top = picIcon(2).Top + picIcon(2).Height + 20
    Label1(3).Top = picIcon(3).Top + picIcon(3).Height + 20
    Label1(4).Top = picIcon(4).Top + picIcon(4).Height + 20
    Label1(5).Top = picIcon(5).Top + picIcon(5).Height + 20
    Label1(6).Top = picIcon(6).Top + picIcon(6).Height + 20
    Label1(7).Top = picIcon(7).Top + picIcon(7).Height + 20
    
    picIcon(1).Top = 60 + picIcon(0).Height + 20 + Label1(0).Height + 20
    picIcon(2).Top = Label1(1).Top + Label1(1).Height + 20
    picIcon(3).Top = Label1(2).Top + Label1(2).Height + 20
    picIcon(4).Top = Label1(3).Top + Label1(3).Height + 20
    picIcon(5).Top = Label1(4).Top + Label1(4).Height + 20
    picIcon(6).Top = Label1(5).Top + Label1(5).Height + 20
    picIcon(7).Top = Label1(6).Top + Label1(6).Height + 20
Label1(b - 1).Caption = FileList.List(b - 1 + IconStart)
Label1(b - 1).Visible = True
picIcon(b - 1).Visible = True
picIcon(b - 1).ToolTipText = FileList.List(b - 1 + IconStart)
GoTo 10
End Sub
Private Sub Drive_Change()
C$ = Drive
On Error GoTo 10
Folder.Path = Drive
DefaultDrive = Drive
DefaultPath = Folder.Path
Exit Sub
10 Drive = DefaultDrive
Folder.Path = DefaultPath
MsgBox "Drive Reading Error", vbOKOnly, "Error!!!"
sndPlay "BOING", SoundOps.SND_ASYNC
End Sub
Private Sub Folder_Change()

frmChart.Cls
 If picOut.Visible = True Then
sndPlay "bloop_x", SoundOps.SND_ASYNC
    picBuffer.Visible = False
    picBackBuffer.Visible = False
    picOut.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    End If
    'IconStart = 0
    'Update_Icons
    ScaleBar.Enabled = True
    Mode = About
    imgPreset.Visible = False
    cmdPreset.Visible = False
    If Str$(FileList.ListCount) > 0 Then
    Picture1.Visible = True
    IconScroll.Visible = True
    
    End If
    
    Picture2.Visible = True
    lblInfo.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
FileList.Path = Folder.Path
DefauultPath = Folder.Path
'mnuCopy.Enabled = True
'mnuPrint.Enabled = True
'mnuSave.Enabled = True
t = FileList.ListCount - 7
If t < 0 Then t = 0
IconScroll.Max = t
IconScroll.Value = 0
FileList.ToolTipText = Str$(FileList.ListCount) + " Items"
Update_Icons
lblInfo.Caption = FileList.ListCount & " Files"
End Sub
Private Sub FileList_Click()
frmChart.Cls
If picOut.Visible = True Then
sndPlay "bloop_x", SoundOps.SND_ASYNC
    picBuffer.Visible = False
    picBackBuffer.Visible = False
    picOut.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    End If
    IconStart = 0
    Update_Icons
    ScaleBar.Enabled = True
    Mode = About
    imgPreset.Visible = False
    cmdPreset.Visible = False
    IconScroll.Visible = True
    Picture1.Visible = True
    Picture2.Visible = True
    lblInfo.Visible = True
    fraQBColors.Visible = False
    frmChart.Label8.Visible = False
    Me.Picture = LoadPicture(App.Path & "/Images/main.jpg")
    VMove.Visible = True
    HMove.Visible = True
    Label13.Visible = True
    ScaleBar.Visible = True
    ScaleMeter.Visible = True
If Right$(FileList.Path, 1) = "\" Then r$ = "" Else r$ = "\"
FilePath = FileList.Path + r$ + FileList.Filename
FileExt = FileList.Filename
Picture_Shown = True
NewImage.Visible = True
Sample.Picture = LoadPicture(FilePath)
NewImage.Picture = LoadPicture(FilePath)
NewImage_Width = NewImage.Width
NewImage_Height = NewImage.Height
ScaleBar.Enabled = True
ScaleBar.Value = 100
ScaleBar.Max = 500
If NewImage_Width * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Width) * 100
If NewImage_Height * (ScaleBar.Max / 100) > 32000 Then ScaleBar.Max = (32000 / NewImage_Height) * 100
cmdExit.Enabled = True
NewImage.Cls
Picture_Update
'W5535 H6135
End Sub
Private Sub picIcon_Click(Index As Integer)
FileList.Selected(Index + IconStart) = True
End Sub


Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgADD.Visible = False
cmdADD.Visible = True
    
End Sub
Private Sub imgExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgExit.Visible = False
cmdExit.Visible = True
    
End Sub
Private Sub imgPreset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgPreset.Visible = False
cmdPreset.Visible = True

End Sub
Private Sub imgCapture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCapture.Visible = False
cmdCapture.Visible = True
End Sub
Private Sub imgHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHide.Visible = False
cmdHide.Visible = True
    
End Sub
Private Sub imgPicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgPicker.Visible = False
cmdPicker.Visible = True
    
End Sub
Private Sub imageLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imageLoad.Visible = False
cmdLoad.Visible = True
    
End Sub
Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSave.Visible = False
cmdSave.Visible = True
    
End Sub
Private Sub imgCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgCopy.Visible = False
cmdCopy.Visible = True
    
End Sub
Private Sub imgPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgPrint.Visible = False
cmdPrint.Visible = True

End Sub
Private Sub imgRect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgRect.Visible = False
cmdRect.Visible = True

End Sub
Private Sub imgHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgHelp.Visible = False
cmdHelp.Visible = True

End Sub
Private Sub txtH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtH.ForeColor = vbRed

End Sub
Private Sub txtS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtS.ForeColor = vbRed

End Sub
Private Sub txtB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtB.ForeColor = vbRed

End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text1.ForeColor = vbRed

End Sub
Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text2.ForeColor = vbRed

End Sub
Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text3.ForeColor = vbRed

End Sub
Private Sub txtCyan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtCyan.ForeColor = vbRed

End Sub
Private Sub txtMagenta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtMagenta.ForeColor = vbRed

End Sub
Private Sub txtYellow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtYellow.ForeColor = vbRed

End Sub
Private Sub txtK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtK.ForeColor = vbRed

End Sub
Private Sub txtColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With txtColor(Index)
        .ForeColor = vbRed
    End With
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtH.ForeColor = &H800000
txtS.ForeColor = &H800000
txtB.ForeColor = &H800000
Text1.ForeColor = &H800000
Text2.ForeColor = &H800000
Text3.ForeColor = &H800000
txtCyan.ForeColor = &H800000
txtMagenta.ForeColor = &H800000
txtYellow.ForeColor = &H800000
txtK.ForeColor = &H800000
End Sub
Private Sub lblPantone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPantone.ForeColor = vbRed
End Sub


Function StrRev(k As String)
Dim X As Integer, sBuff As String
    'Build a string backwards eg StrRev("ben")'returns neb
    For X = Len(k) To 1 Step -1
        sBuff = sBuff & Mid(k, X, 1)
    Next
    
    X = 0
    StrRev = sBuff
    
End Function

