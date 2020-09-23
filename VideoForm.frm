VERSION 5.00
Begin VB.Form VideoForm 
   Caption         =   " Video File"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   180
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the address of the Video file"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "VideoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub command1_Click()
'Simple, user enters URL of the video, then the URL
'is entered into a string, allong with the correct
'HTML

VideoCode = "<embed src=" & Text1.text & " quality=high ></embed>"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
Call command1_Click
End If
End Sub
