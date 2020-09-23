VERSION 5.00
Begin VB.Form BGSound 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Sound File"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Options"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "All the time (infinite)"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   4095
         Begin VB.OptionButton Option6 
            Caption         =   "6"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3360
            TabIndex        =   10
            Top             =   480
            Width           =   495
         End
         Begin VB.OptionButton Option5 
            Caption         =   "5"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            TabIndex        =   9
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "4"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2160
            TabIndex        =   8
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "3"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1560
            TabIndex        =   7
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "2"
            Enabled         =   0   'False
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Select how many times the file should play"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   3615
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "How many times do you want the file to play?"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the address of the sound file"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "BGSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'Simple validation, if a user chooses to run the sound
'infinite times, then there is no need to ask him how
'many times, so we disable the options buttons.

If Check1.Value = 1 Then
Option1.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Label3.Visible = False
End If

If Check1.Value = 0 Then
Option1.Enabled = True
Option1.Value = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Label3.Visible = True
End If


End Sub

Private Sub command1_Click()
'Makes sure the user has entered text'
If Text1.text = "" Then
MsgBox "You must enter an address for the sound file", vbCritical, "Error"
Exit Sub
End If
''''''''''''''''''''''''''''''''''''

Dim fileplay As String

'Determines how many time the user wants to play the
'File
If Option1.Value = True Then
fileplay = "1"
End If
If Option2.Value = True Then
fileplay = "2"
End If
If Option3.Value = True Then
fileplay = "3"
End If
If Option4.Value = True Then
fileplay = "4"
End If
If Option5.Value = True Then
fileplay = "5"
End If
If Option6.Value = True Then
fileplay = "6"
End If
If Check1.Value = 1 Then
fileplay = "infinite"
End If
''''''''''''''''''''''''''''

'''Input the HTML into a string, from what
'''we have obtained from the user
BGSoundCode = "<bgsound src=" & Chr(34) & Text1.text & Chr(34) & " loop=" & Chr(34) & fileplay & Chr(34) & ">"
''''''''''''''''''''

Unload Me 'Close the form
End Sub

Private Sub Command2_Click()
Unload Me 'Close the form, if the user hits cancel
End Sub

