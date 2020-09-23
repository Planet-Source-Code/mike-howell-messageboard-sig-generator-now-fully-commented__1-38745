VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Linkform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Hyperlink"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Align"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   5775
      Begin VB.OptionButton Option12 
         Caption         =   "Right"
         Height          =   255
         Left            =   3960
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Left"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Center"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Colour"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "Select Font Colour"
         Height          =   465
         Left            =   480
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "#000000"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   5775
      Begin VB.OptionButton Option10 
         Caption         =   "7"
         Height          =   225
         Left            =   4320
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option9 
         Caption         =   "6"
         Height          =   225
         Left            =   3720
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option8 
         Caption         =   "5"
         Height          =   225
         Left            =   3120
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option7 
         Caption         =   "4"
         Height          =   225
         Left            =   4920
         TabIndex        =   18
         Top             =   600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option6 
         Caption         =   "3"
         Height          =   225
         Left            =   4320
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "2"
         Height          =   225
         Left            =   3720
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "1"
         Height          =   225
         Left            =   3120
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Select Font Type"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cosmic Sans MS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Link Type"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton Option3 
         Caption         =   "Ftp"
         Height          =   345
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "E-mail"
         Height          =   345
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Website"
         Height          =   345
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5775
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Enter what you want the link to say. If you leave it blank then it will be the same as the target address:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "http://"
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter website address:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "Linkform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

'Set our variables..

Dim Linktype As String
Dim WebsiteAddress As String
Dim LinkText As String
Dim Fonttype As String
Dim FontSize As String
Dim FontColor As String
Dim LinkAlign As String


'Validate, make sure the user entereed text
If Text2.text = "" Then
Text2.text = Text1.text
End If

WebsiteAddress = Text1.text 'write the website into the string

'Check the options values to see what type of link
'You want. (mail, url, ftp)
If Option2.Value = True Then
WebsiteAddress = "mailto:" & WebsiteAddress
End If

If Option3.Value = True Then
WebsiteAddress = "ftp://" & WebsiteAddress
End If
'''''''''''''''''''''''''''''''''''''''''''''''''


LinkText = Text2.text 'write the text the user sets into the string

Fonttype = Label3.Caption 'Get the font Name and write it into the string
Fonttype = Trim(Fonttype) 'Trim the font name, so no spaces.

'''''''''''Checks what the user selects and set
'''''''''''the font size accordingly
If Option4.Value = True Then
FontSize = "1"
End If

If Option5.Value = True Then
FontSize = "2"
End If

If Option6.Value = True Then
FontSize = "3"
End If

If Option7.Value = True Then
FontSize = "4"
End If

If Option8.Value = True Then
FontSize = "5"
End If

If Option9.Value = True Then
FontSize = "6"
End If

If Option10.Value = True Then
FontSize = "7"
End If
''''''''''''''''''''''''''''''''''''''''''

FontColor = HexCode 'Load the colour code into the string

''Write the URL HTML depending on what the user selects
LinkHTMLCode = "<font face=" & FontName & " color=" & FontColor & " size=" & FontSize & " >" & "<a href=" & Chr(34) & WebsiteAddress & Chr(34) & ">" & LinkText & "</a></font>"



''After we have written the code, then we add our
''Alignment code
If Option11.Value = True Then
LinkHTMLCode = "<center>" & LinkHTMLCode & "</center>"
End If

If Option12.Value = True Then
LinkHTMLCode = "<div align=" & Chr(34) & "left" & Chr(34) & ">" & LinkHTMLCode & "</div>"
End If

If Option13.Value = True Then
LinkAlign = "<div align=" & Chr(34) & "right" & Chr(34) & ">" & LinkHTMLCode & "</div>"
End If
'''''''

Unload Me
End Sub

Private Sub Command3_Click()
FontLists.Show 1 'Show the font form
Label3.Caption = strFontName
Label3.FontName = strFontName
strFontName = ""
End Sub

Private Sub Command4_Click()
HexCode = ""

ColorForm.Show 1

Label5.Caption = HexCode
End Sub

Private Sub Option1_Click()
Label1.Caption = "Enter website address:"
Text1.text = "http://"
End Sub

Private Sub Option2_Click()
Label1.Caption = "Enter e-mail address:"
Text1.text = ""
End Sub

Private Sub Option3_Click()
Label1.Caption = "Enter ftp server:"
Text1.text = ""
End Sub
