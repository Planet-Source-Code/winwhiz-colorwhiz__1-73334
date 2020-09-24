VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " About ColorWhiz"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":076A
   ScaleHeight     =   4860
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ReDrawTimer 
      Interval        =   1
      Left            =   7620
      Top             =   5760
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
      Height          =   4215
      Left            =   113
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   120
      Width           =   6495
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
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
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
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   2948
      Picture         =   "frmAbout.frx":2166
      Top             =   4440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2948
      Picture         =   "frmAbout.frx":25AB
      Top             =   4440
      Width           =   825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
   As Long

Private Const SRCCOPY = &HCC0020

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
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image1.Visible = False
Image2.Visible = True
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image1.Visible = True
Image2.Visible = False
    
End Sub
Private Sub Form_Load()
sndPlay "xfilesintro", SoundOps.SND_ASYNC
Dim iLine As Integer

    NumLines = 1
    
    frmAbout.ScaleMode = vbPixels
    
    picBuffer.ScaleMode = vbPixels
    
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    
    picBuffer.Visible = False
    
    Open (App.Path & "\credits.txt") For Input As #1
    
    Do Until EOF(1)
        Line Input #1, Tempstring(NumLines)
        NumLines = NumLines + 1
    Loop
    Close #1
    
    NumLines = NumLines - 1
    
    lX = picBuffer.ScaleLeft
    lY = picBuffer.ScaleHeight
    
    GradiantBackground picBackBuffer
    
    ReDrawTimer.Interval = 19
    ReDrawTimer.Enabled = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
    sndPlay "bloop_x", SoundOps.SND_ASYNC
    frmChart.Visible = True
    Unload Me
    
End Sub

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

Private Sub Image2_Click()
ReDrawTimer.Enabled = False
sndPlay "bloop_x", SoundOps.SND_ASYNC
frmChart.Visible = True
    Unload Me
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
                
                'If j = NumLines And picBuffer.CurrentY < -25 Then
                'If we've painted the last line, and it's above the top, there's no more text to scroll
                ' and we exit.
                    'ReDrawTimer.Enabled = False
                    'sndPlay "bloop_x", SoundOps.SND_ASYNC
                    'frmChart.Visible = True
                    'Unload Me
                'End If
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



