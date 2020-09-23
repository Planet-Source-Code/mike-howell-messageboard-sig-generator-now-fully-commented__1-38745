VERSION 5.00
Begin VB.Form hrline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Horizontal Line"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
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
      Left            =   3600
      TabIndex        =   11
      Top             =   3360
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alignment"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4815
      Begin VB.OptionButton Option3 
         Caption         =   "Right"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Left"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line properties"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox list1 
         Height          =   315
         ItemData        =   "hrline.frx":0000
         Left            =   2760
         List            =   "hrline.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Select this thickness you want the line to be:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the percentage of the page you want the line to cover:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "hrline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
''Validation: Make sure the user has entered data
If Text1.text = "" Then
MsgBox "You must select the line width", vbCritical, "Error"
End If

If list1.text = "" Then
MsgBox "You must select the line thickness", vbCritical, "Error"
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'3 Different routines, each one dependends on the option the user selected'
If Option1.Value = True Then
HRLineCode = "<hr width=" & Chr(34) & Text1.text & "%" & Chr(34) & " size=" & list1.text & " >" 'Load the HTML code into the string
End If

If Option2.Value = True Then
HRLineCode = "<hr width=" & Chr(34) & Text1.text & "%" & Chr(34) & " size=" & list1.text & " align=" & Chr(34) & "left" & Chr(34) & " >" 'Load the HTML code into the string
End If

If Option3.Value = True Then
HRLineCode = "<hr width=" & Chr(34) & Text1.text & "%" & Chr(34) & " size=" & list1.text & " align=" & Chr(34) & "right" & Chr(34) & " >" 'Load the HTML code into the string
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''
Unload Me
End Sub

Private Sub Form_Load()
Dim listitem As Integer

For listitem = 1 To 10 'Lists 1 to 10 in a list box, used for the font size
list1.AddItem listitem
Next listitem

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
'Make sure only numbers can be entered into text1
KeyAscii = KeyAscii * Abs(((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = vbKeyBack))
End Sub
