VERSION 5.00
Begin VB.MDIForm mdi_Help 
   BackColor       =   &H00000000&
   Caption         =   " ColorWhiz Help"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "mdi_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MsgBox "Click on a picture of a control to learn about it's use.", vbInformation, "ColorWhiz Help"
frmHelp.Show
End Sub
