VERSION 5.00
Begin VB.Form Advanced 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Advanced Controls"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tag Information"
      Height          =   375
      Left            =   2040
      TabIndex        =   26
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label maindisplay 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Copyright 
      BackColor       =   &H80000001&
      Caption         =   "            ©"
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
      Left            =   1800
      TabIndex        =   25
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Header 
      BackColor       =   &H80000001&
      Caption         =   "    <h*> </h*>"
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
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Reg 
      BackColor       =   &H80000001&
      Caption         =   "            ®"
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
      Left            =   1800
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Frame 
      BackColor       =   &H80000001&
      Caption         =   "  <frame> </frame>"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Superscript 
      BackColor       =   &H80000001&
      Caption         =   "     <sup> </sup>"
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
      Left            =   1800
      TabIndex        =   21
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblFont 
      BackColor       =   &H80000001&
      Caption         =   "   <font> </font>"
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
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Subscript 
      BackColor       =   &H80000001&
      Caption         =   "     <sub> </sub>"
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
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Center 
      BackColor       =   &H80000001&
      Caption         =   "   <center></center>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Strong 
      BackColor       =   &H80000001&
      Caption         =   " <strong></strong>"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label LineBreak 
      BackColor       =   &H80000001&
      Caption         =   "       <br>"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Strike 
      BackColor       =   &H80000001&
      Caption         =   "  <strike></strike>"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Body 
      BackColor       =   &H80000001&
      Caption         =   "   <body> </body>"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label script 
      BackColor       =   &H80000001&
      Caption         =   "  <script></script>"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Blink 
      BackColor       =   &H80000001&
      Caption         =   "   <blink> </blink>"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Paragraph 
      BackColor       =   &H80000001&
      Caption         =   "     <p> </p>"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Big 
      BackColor       =   &H80000001&
      Caption         =   "    <big> </big>"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Marquee 
      BackColor       =   &H80000001&
      Caption         =   "    <marquee>"
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
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Bold 
      BackColor       =   &H80000001&
      Caption         =   "      <b> </b>"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblImage 
      BackColor       =   &H80000001&
      Caption         =   "       <img>"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Anchor 
      BackColor       =   &H80000001&
      Caption         =   "      <a> </a>"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Italics 
      BackColor       =   &H80000001&
      Caption         =   "      <i> </i>"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Comment 
      BackColor       =   &H80000001&
      Caption         =   "      <!-- -->"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label49 
      BackColor       =   &H80000001&
      Caption         =   "Label2"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label48 
      BackColor       =   &H80000001&
      Caption         =   "Label2"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label44 
      BackColor       =   &H80000001&
      Caption         =   "Label2"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label43 
      BackColor       =   &H80000001&
      Caption         =   "Label2"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "Advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All the code below, pretty much is just for
'changing the colours of things, as you may see
'Look for command1_click for the main code. It just
'opens a webpage to the section of the site
'that contains information about the tag the user
'selected

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Anchor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

With Anchor
.BackColor = &H80000002
.ForeColor = &H80000001
End With

End Sub

Private Sub Anchor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Anchor
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Anchor.Caption
End Sub


Private Sub Big_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Big
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Big_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Big
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Big.Caption
End Sub


Private Sub Blink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Blink
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Blink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With Blink
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Blink.Caption
End Sub


Private Sub Body_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Body
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Body_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Body
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Body.Caption
End Sub


Private Sub Bold_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Bold
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Bold_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Bold
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Bold.Caption
End Sub


Private Sub Center_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Center
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Center_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Center
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Center.Caption
End Sub

Private Sub command1_Click()
Dim IEOpen As Long

Select Case maindisplay.Caption

Case "      <!-- -->"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Comment", vbNullString, vbNullString, SW_MAXIMIZE)

Case "      <a> </a>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Anchor", vbNullString, vbNullString, SW_MAXIMIZE)

Case "      <b> </b>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Bold", vbNullString, vbNullString, SW_MAXIMIZE)

Case "    <big> </big>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Big%20Text", vbNullString, vbNullString, SW_MAXIMIZE)

Case "   <blink> </blink>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Blink", vbNullString, vbNullString, SW_MAXIMIZE)


Case "   <body> </body>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/toplevel.html#Body", vbNullString, vbNullString, SW_MAXIMIZE)

Case "       <br>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/pformat.html#Line%20Break", vbNullString, vbNullString, SW_MAXIMIZE)

Case "   <center></center>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/pformat.html#Center", vbNullString, vbNullString, SW_MAXIMIZE)

Case "   <font> </font>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Font", vbNullString, vbNullString, SW_MAXIMIZE)

Case "  <frame> </frame>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Frame", vbNullString, vbNullString, SW_MAXIMIZE)

Case "    <h*> </h*>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/pformat.html#Heading%201", vbNullString, vbNullString, SW_MAXIMIZE)

Case "      <i> </i>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Italic", vbNullString, vbNullString, SW_MAXIMIZE)

Case "       <img>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Inline%20Image", vbNullString, vbNullString, SW_MAXIMIZE)

Case "    <marquee>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Marquee", vbNullString, vbNullString, SW_MAXIMIZE)

Case "     <p> </p>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/pformat.html#Paragraph", vbNullString, vbNullString, SW_MAXIMIZE)

Case "  <script></script>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/other.html#Script", vbNullString, vbNullString, SW_MAXIMIZE)

Case "  <strike></strike>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Strikethrough", vbNullString, vbNullString, SW_MAXIMIZE)

Case " <strong></strong>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Strong", vbNullString, vbNullString, SW_MAXIMIZE)

Case "     <sub> </sub>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Subscript", vbNullString, vbNullString, SW_MAXIMIZE)

Case "     <sup> </sup>"
IEOpen = ShellExecute(Me.hwnd, "open", "http://www.willcam.com/cmat/html/lformat.html#Superscript", vbNullString, vbNullString, SW_MAXIMIZE)

Case "            ®"
MsgBox "Type in &reg to achieve this", vbInformation, "Registered"
Case "            ©"
MsgBox "Type on &copy to achieve this", vbInformation, "Copyright"
End Select

End Sub

Private Sub Command2_Click()
Unload Me
Main.Show
End Sub

Private Sub Comment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

With Comment
.BackColor = &H80000002
.ForeColor = &H80000001
End With

End Sub

Private Sub Comment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

With Comment
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Comment.Caption
End Sub

Private Sub Copyright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Copyright
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Copyright_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Copyright
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Copyright.Caption
End Sub

Private Sub Frame_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Frame
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Frame.Caption
End Sub

Private Sub Header_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Header
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Header.Caption
End Sub


Private Sub lblImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lblImage
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = lblImage.Caption
End Sub

Private Sub Italics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Italics
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Italics.Caption
End Sub

Private Sub lblFont_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lblFont
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Frame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Frame
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Header_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Header
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub



Private Sub lblImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lblImage
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Italics_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Italics
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub


Private Sub lblFont_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lblFont
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = lblFont.Caption
End Sub

Private Sub LineBreak_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With LineBreak
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub


Private Sub LineBreak_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With LineBreak
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = LineBreak.Caption
End Sub

Private Sub Marquee_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Marquee
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub


Private Sub Marquee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Marquee
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Marquee.Caption
End Sub

Private Sub Paragraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Paragraph
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub


Private Sub Paragraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Paragraph
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Paragraph.Caption
End Sub

Private Sub Reg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Reg
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub



Private Sub Reg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Reg
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Reg.Caption
End Sub

Private Sub script_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With script
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub script_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With script
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = script.Caption
End Sub

Private Sub Strike_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Strike
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Strike_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Strike
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Strike.Caption
End Sub

Private Sub Strong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Strong
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Strong_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Strong
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Strong.Caption
End Sub

Private Sub Subscript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Subscript
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub


Private Sub Subscript_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Subscript
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Subscript.Caption
End Sub

Private Sub Superscript_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Superscript
.BackColor = &H80000002
.ForeColor = &H80000001
End With
End Sub

Private Sub Superscript_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Superscript
.BackColor = &H80000001
.ForeColor = &H80000012
End With
maindisplay.Caption = Superscript.Caption
End Sub
