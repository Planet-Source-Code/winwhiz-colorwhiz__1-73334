VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const mcsClassName As String = "frmPrint"

'***************************************************************
''--- 12/12/2001 GCC - Windows API/Global Declarations for :PrintScreenUsingKeyboardEvent
'***************************************************************
Private Declare Sub keybd_event Lib "User32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)

Private Const TheScreen = 1
Private Const TheForm = 0

Public Sub PrintScreenSnapShot()
'--- 12/12/2001 GCC - Added print screen feature
'--- Call this from any form in your project.
'--- Example: To make "F6" key a print screen button of a form in your project,
'--- set the form's .KeyPreview property to "True" at design time, and add an
'--- event handler for the form's KeyDown event that looks something like this:
'
'    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'        'Enable PrintScreen feature
'        '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
'        If Shift = 0 Then
'            Select Case KeyCode
'                Case vbKeyF6
'                    KeyCode = 0
'                    PrintScreenSnapShot
'            End Select
'        End If
'    End Sub
'
'--- NOTE: I haven't fully finished tweaking this routine for handling different types of
'--- clipboard contents. Right now, if you just have plain text, Bitmap, or RTF text in the
'--- Windows clipboard, it will cache, and restore the contents. This is necessary because
'--- this PrintScreen features uses the clipboard to handle the PrintScreen image. Because
'--- I put in an "On Error Resume Next" the print screen should work, but the Windows
'--- clipboard contents may get lost if it contains formats other than plain text, RTF text, or
'--- Bitmaps.  The code to manage the clipboard data was copied from VB Help.

    Const csMethodName As String = "PrintScreenSnapShot"
    Dim ClpFmt As Variant
    Dim vContents As Variant

    On Error GoTo ErrorHandler

    If PrinterIsInstalled Then
        On Error Resume Next    ' Set up error handling.
        If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
        If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
        If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
        If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
        
        'On Error GoTo ErrorHandler
        
        '--- Cache current contents of clipboard:
        Select Case ClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                vContents = Clipboard.GetText(vbCFText)
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                '--- 03/19/2002 GCC - Use "Set" in this case and drop optional param:
                Set vContents = Clipboard.GetData
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                '--- 03/19/2002 GCC - ...
                '--- Not sure if this is correct because I'm not sure how
                '--- to set both text and a bitmap into the clipboard
                vContents = Clipboard.GetData
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                vContents = Clipboard.GetText(vbCFRTF)
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
        'MsgBox Msg  ' Display message.
        
        '--- Do a <Print Scrn> with an API call:
        keybd_event vbKeySnapshot, TheForm, 0&, 0&

        '--- Give Windows a chance to update the clipboard
        DoEvents
        
        PrintBitmap Clipboard.GetData(vbCFBitmap)  ', Me.Height, Me.Width
        
        '--- Restore contents of clipboard
        Clipboard.Clear
        Select Case ClpFmt
            Case 1
                'Msg = "The Clipboard contains only text."
                Clipboard.SetText vContents, vbCFText
            Case 2, 4, 6
                'Msg = "The Clipboard contains only a bitmap."
                '--- 03/19/2002 GCC - Drop optional param:
                 'Clipboard.SetData vContents, ClpFmt
                 Clipboard.SetData vContents
            Case 3, 5, 7
                'Msg = "The Clipboard contains text and a bitmap."
                '--- 03/19/2002 GCC - ....
                '--- Not sure if this is correct because I'm not sure how
                '--- to set both text and a bitmap into the clipboard
                'Clipboard.SetData vContents, ClpFmt
                Clipboard.SetData vContents
            Case 8, 9
                'Msg = "The Clipboard contains only rich text."
                '--- Example: Copied text inside MSWord
                Clipboard.SetText vContents, vbCFRTF
            Case Else
                'Msg = "There is nothing on the Clipboard."
        End Select
    End If

    Exit Sub

ErrorHandler:
End Sub

Private Function PrinterIsInstalled() As Boolean
'--- 12/12/2001 GCC - added to support print screen feature
    Dim dummy As String

    On Error Resume Next
    dummy = Printer.DeviceName
    
    If Err.Number Then
        sndPlay "BOING", SoundOps.SND_ASYNC
        MsgBox "No default printer installed." & vbCrLf _
            & "To install and select a default printer, select the " _
            & "Setting / Printers command in the Start menu, and then " _
            & "double-click on the Add Printer icon.", _
            vbExclamation, "Printer Error"
        PrinterIsInstalled = False
    Else
        PrinterIsInstalled = True
    End If
End Function

Public Sub PrintBitmap(picBitmap As Variant)
    Const csMethodName As String = "PrintBitmap"

    On Error GoTo ErrorHandler

    With Me
        .Visible = False
        .Height = Screen.Height + 300
        .Width = Screen.Width + 150
        .Picture = picBitmap
        .PrintForm
    End With
    Unload Me

    Exit Sub

ErrorHandler:
End Sub


