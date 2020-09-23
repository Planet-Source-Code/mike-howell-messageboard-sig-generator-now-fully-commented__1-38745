VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CodeWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Window"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ForeColor       =   &H00000000&
   Icon            =   "CodeWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Info"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   195
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Code To Clipboard"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Redirect Page"
      Top             =   120
      Width           =   3855
   End
   Begin RichTextLib.RichTextBox RichTxtBox 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"CodeWindow.frx":1042
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CodeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'if the user hits enter, then make the the program
'run the code, in Command3_click

If KeyAscii = "13" Then
Call Command3_Click
End If
End Sub

Private Sub command1_Click()
Unload Me
Main.Show
End Sub

Private Sub Command2_Click()
Clipboard.SetText RichTxtBox.TextRTF 'Copy the text in the Rich text box
End Sub

Private Sub Command3_Click()
Select Case Combo1.text

Case "Redirect Page"
RichTxtBox.LoadFile App.Path & "\Scripts\Redirect.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Automatic Refresh"
RichTxtBox.LoadFile App.Path & "\Scripts\Refresh.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Add Time to Page"
RichTxtBox.LoadFile App.Path & "\Scripts\DigAnClock.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Disable Right Click"
RichTxtBox.LoadFile App.Path & "\Scripts\DisableRC.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Change Status Bar Message"
RichTxtBox.LoadFile App.Path & "\Scripts\StatusBar.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Add Current Date"
RichTxtBox.LoadFile App.Path & "\Scripts\CurrentDate.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

Case "Add Page/Board to Bookmarks"
RichTxtBox.LoadFile App.Path & "\Scripts\Bookmark.txt"
    ' Set the RichTextBox for the Color coding Control
    ' This is needed because of the way the control does it's color coding
     'EZColorCode.RichTxtBox = RichTxtBox
    
    ' Set the colors:
    m_TextCol = vbBlack
    m_AttribCol = 8388736
    m_TagCol = 10485760
    m_CommentCol = 8421440
    m_AspCol = 128
    
    ' Now lets add the template and color code it. First hide the text box from the user
    RichTxtBox.Visible = True
    RichTxtBox.AutoVerbMenu = True
    RichTxtBox.HideSelection = True
HtmlHighlight

End Select

End Sub

Private Sub Command4_Click()
MsgBox "This is not code than can be used in Most Messageboard Sigs, but can be used by board admins or webmasters. More code will be added in further releases", vbInformation, "Information" 'Little info
End Sub

Private Sub Form_Load()
'Add our options to our combo box, easy for the user
'to select

Setcolors
Combo1.AddItem "Redirect Page"
Combo1.AddItem "Automatic Refresh"
Combo1.AddItem "Add Time to Page"
Combo1.AddItem "Disable Right Click"
Combo1.AddItem "Change Status Bar Message"
Combo1.AddItem "Add Current Date"
Combo1.AddItem "Add Page/Board to Bookmarks"
End Sub

Private Sub Rich1_KeyUp(KeyCode As Integer, Shift As Integer)
Setcolors
End Sub

Private Sub Rich1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Setcolors
End Sub


Public Sub Setcolors()
'set the colours for the HTML


commentchar = "'"
longvar = "$"
If KeyCode = 13 Then Exit Sub
LockWindowUpdate Me.hwnd
clearwordcolors RichTxtBox
ColorizeWord RichTxtBox, longvar, &H80& 'This is for vars with no fixed lenght e.g (in perl _
it could be $122434 0r $myvarx or $x ..... Always set this ($) as first word to colorize
ColorizeWord RichTxtBox, commentchar, &H8000& 'This char is for comments like this
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ColorizeWord RichTxtBox, "<!-- Insert Content Here -->", &HDE3E3
ColorizeWord RichTxtBox, "<SCRIPT", &H80FF&
LockWindowUpdate 0&
RichTxtBox.Enabled = True
If RichTxtBox.Visible = True Then
RichTxtBox.SetFocus
End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub
