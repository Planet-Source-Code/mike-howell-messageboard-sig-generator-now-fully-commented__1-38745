VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Textform 
   Caption         =   " Text"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Effects"
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   5775
      Begin VB.CheckBox Check6 
         Caption         =   "Superscript"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Subscript"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Strike-through"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Font"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Select Font Type"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "1"
         Height          =   225
         Left            =   3120
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option5 
         Caption         =   "2"
         Height          =   225
         Left            =   3720
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option6 
         Caption         =   "3"
         Height          =   225
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option7 
         Caption         =   "4"
         Height          =   225
         Left            =   4920
         TabIndex        =   9
         Top             =   600
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option8 
         Caption         =   "5"
         Height          =   225
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option9 
         Caption         =   "6"
         Height          =   225
         Left            =   3720
         TabIndex        =   7
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option10 
         Caption         =   "7"
         Height          =   225
         Left            =   4320
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   1320
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
         TabIndex        =   15
         Top             =   840
         Width           =   1815
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
   End
   Begin VB.Frame Frame4 
      Caption         =   "Colour"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "Select Font Colour"
         Height          =   465
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "#000000"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Enter what you want the text to say:"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "textform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim text As String
Dim Textsize As String
Dim Fonttype As String
Dim FontColor As String
Dim temp As String

text = Text1.text



If Option4.Value = True Then
Textsize = "1"
End If
If Option5.Value = True Then
Textsize = "2"
End If
If Option6.Value = True Then
Textsize = "3"
End If
If Option7.Value = True Then
Textsize = "4"
End If
If Option8.Value = True Then
Textsize = "5"
End If
If Option9.Value = True Then
Textsize = "6"
End If
If Option10.Value = True Then
Textsize = "7"
End If

Fonttype = Label3.Caption
Fonttype = Trim(Fonttype)

FontColor = Label5.Caption

temp = "<font size=" & Textsize & " type=" & Fonttype & " color=" & FontColor & " >" & text & "</font>"

If Check1.Value = 1 Then
temp = "<b>" & temp & "</b>"
End If

If Check2.Value = 1 Then
temp = "<i>" & temp & "</i>"
End If

If Check3.Value = 1 Then
temp = "<u>" & temp & "</u>"
End If

If Check4.Value = 1 Then
temp = "<strike>" & temp & "</strike>"
End If

If Check5.Value = 1 Then
temp = "<sub>" & temp & "</sub>"
End If

If Check6.Value = 1 Then
temp = "<sup>" & temp & "</sup>"
End If

FontHTMLCode = temp

Unload Me
End Sub

Private Sub Command3_Click()
FontLists.Show 1
Label3.Caption = strFontName
Label3.FontName = strFontName
strFontName = ""
End Sub

Private Sub Command4_Click()
HexCode = ""

ColorForm.Show 1

Label5.Caption = HexCode
End Sub

