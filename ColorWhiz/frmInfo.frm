VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   0  'None
   Caption         =   " ColorWhiz Color Links"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInfo.frx":076A
   ScaleHeight     =   4695
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Color Matters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1620
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Management and Color Science"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1620
      TabIndex        =   3
      Top             =   2100
      Width           =   4020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Psychology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1620
      TabIndex        =   2
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web Color Overview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1620
      TabIndex        =   1
      Top             =   1260
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Analysis and Effects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Left            =   1620
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API call to execute commands with the windows shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
DisplayAsURL Label1
DisplayAsURL Label2
DisplayAsURL Label3
DisplayAsURL Label4
DisplayAsURL Label5
Me.Left = 2800
Me.Top = 300
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0FFFF
Label2.ForeColor = &HC0FFFF
Label3.ForeColor = &HC0FFFF
Label4.ForeColor = &HC0FFFF
Label5.ForeColor = &HC0FFFF
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.ForeColor = &HC0FFC0

End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label2.ForeColor = &HC0FFC0

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label3.ForeColor = &HC0FFC0

End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label4.ForeColor = &HC0FFC0

End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label5.ForeColor = &HC0FFC0

End Sub
Private Sub Label1_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.wwwearables.com/techniques/color.htm"
End Sub

Private Sub Label2_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.webcolors.freeserve.co.uk/colorinfo.htm"
End Sub
Private Sub Label3_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.infoplease.com/spot/colors1.html"
End Sub
Private Sub Label4_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.normankoren.com/color_management.html"
End Sub
Private Sub Label5_Click()
    ' user clicked on lable, lets open the web browser
    BrowseTo "http://www.colormatters.com/comput.html"
End Sub
Private Sub BrowseTo(ByRef pstrURL As String)
    ' Opens users default web browser and navigates to the selected URL
    Call ShellExecute(Me.hwnd, "Open", pstrURL, "", "", True)
End Sub

Private Sub DisplayAsURL(ByRef Link As VB.Label)
    ' Changes a link to look like a URL
    Link.Font.Underline = True
    Link.ForeColor = vbBlue
    Link.MousePointer = vbCustom
    Link.MouseIcon = LoadPicture(App.Path & "\Hand.cur")
End Sub

