VERSION 5.00
Begin VB.Form ImageForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Image"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Image Alignment"
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
      TabIndex        =   11
      Top             =   4200
      Width           =   5415
      Begin VB.OptionButton Option3 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Image Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   5415
      Begin VB.TextBox Text4 
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
         Left            =   2520
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text3 
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Alternative Text: This is displayed when the image does not load"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Border thinckness: (0-50) 0 is no border"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   4
      Text            =   "http://"
      Top             =   120
      Width           =   3135
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
      Left            =   2760
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
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
      Left            =   4200
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Link"
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
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000A&
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
         Left            =   2280
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add a link to the image"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Target Website:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Image Address:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   160
      Width           =   2415
   End
End
Attribute VB_Name = "ImageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
'if the user wants to link an image, then we enable
' the textbox so the user can enter a url
If Check1.Value = 1 Then
Text2.BackColor = &H80000005 'Changes backcolour to white
Text2.Enabled = True 'Disable the textbox
Text2.text = "http://" 'Enters http:// into to box, so the user knows to enter a link
End If

'If the user selects not to link an image then
'we disable the textbox
If Check1.Value = 0 Then
Text2.BackColor = &H8000000A 'Change backcolour to grey
Text2.text = "" 'Clear the text box
Text2.Enabled = False 'Disable the textbox
End If
End Sub

Private Sub command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
'Dim out variables, we will use these to load the options
'the user selects into a string
Dim ImageAddress As String
Dim Imagelink As String
Dim Imageborder As Integer
Dim AlternativeText As String
Dim ImageAlignment As String
''''''''''''''''''''''''''''''

'Valdation, make sure the user has entered text
If Text3.text = "" Then
MsgBox "Please specify an image border size", vbExclamation, "Image Border"
Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''

ImageAddress = Text1.text 'Load the image URL into a string


If Check1.Value = 1 Then 'If the user selects to add a URL to an image
Imagelink = Text2.text 'write the URL into a string
End If

Imageborder = Text3.text 'write the image border into an integer

AlternativeText = Text4.text 'write the alt text into the string

If Option1.Value = True Then 'check the option
ImageAlignment = "center" 'Write the image alignment
End If

If Option2.Value = True Then 'check the option
ImageAlignment = "left" 'Write the image alignment
End If

If Option3.Value = True Then 'check the option
ImageAlignment = "right" 'Write the image alignment
End If

'''2 Routines, one is carried out depending on the user settings'
If Check1.Value = 1 Then
ImageHTMLCode = "<a href=" & Chr(34) & Imagelink & Chr(34) & "><img src=" & Chr(34) & ImageAddress & Chr(34) & " border=" & Imageborder & " align=" & ImageAlignment & " alt=" & AlternativeText & " ></a>"
End If

If Check1.Value = 0 Then
ImageHTMLCode = "<img src=" & Chr(34) & ImageAddress & Chr(34) & " border=" & Imageborder & " align=" & ImageAlignment & " alt=" & AlternativeText & " >"
End If
'''''''''''''''''''''''''''''''''''''''''''''

Unload Me
End Sub

Private Sub Command3_Click()
MsgBox "This section is still under construction", vbExclamation, "Coming Soon.."
Exit Sub
Text1.SetFocus
frmGetPics.Show 1
End Sub

Private Sub Form_Load()
'set the option1 to be clicked when the form loads
Option1.Value = True

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'Make sure that only numbers can be entered
    'Into the text3 box
    
    Dim Numbers As Integer
    Dim Msg As String
    Numbers = KeyAscii


    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        KeyAscii = 0
    End If
End Sub
