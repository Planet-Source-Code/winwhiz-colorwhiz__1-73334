VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1890
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   2100
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal _
    hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
   lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Const PS_SOLID = 0
 
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    
    
    
End Sub

Private Sub Form_Load()
    Dim StartTime As Long, s As String, sTimer As Long
    Me.Picture1.Picture = LoadPicture(App.Path & "\Wiz_01.jpg")
    Me.Picture1.Refresh
    Me.Show
    StartTime = timeGetTime
    Laser Picture1
    s = Format((timeGetTime - StartTime) / 1000, "0.00") & " seconds"
    Debug.Print "LineTo API: " & s
End Sub

Private Function Laser(PicBox As PictureBox) '

    Dim Color As Long, i As Long, j As Long, SrcDC As Long, DstDC As Long
    Dim StartX As Long, StartY As Long, DestX As Long, DestY As Long
    Dim LastColor As Long, hPen As Long
    
    With PicBox
         SrcDC = .hdc
         StartX = (frmIntro.ScaleWidth - PicBox.ScaleWidth) / 2
         StartY = (frmIntro.ScaleHeight - PicBox.ScaleHeight) / 2
         DestX = PicBox.ScaleWidth
         DestY = PicBox.ScaleHeight
    End With
    DstDC = frmIntro.hdc
    DoEvents
    
    For j = 0 To DestX - 1
        For i = 0 To DestY - 1
            Color = GetPixel(SrcDC, j, i)
            If LastColor <> Color Then
                'don 't change the pen color unless necessary! slows it down.
                hPen = CreatePen(PS_SOLID, 1, Color): DeleteObject SelectObject(DstDC, hPen)
                LastColor = Color
            End If
            MoveToEx DstDC, j + StartX, i + StartY, ByVal 0&
            LineTo DstDC, DestX + StartX, DestY + StartY
        Next i
    Next j
    
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    
    
    
    End Sub
