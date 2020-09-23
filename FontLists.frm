VERSION 5.00
Begin VB.Form FontLists 
   Caption         =   " Font"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FontLists.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4440
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1333 
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
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FontLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Click()
    Label1.FontName = Combo1.text 'Change the label font, to the font we select
End Sub


Private Sub command1_Click()
strFontName = Label1.FontName 'Change the label to the font name
Unload Me
End Sub

Private Sub Form_Load()

'Load every font into the combo box
'A little slow, but my Common Dialog wouldent work
'So i had to use an alternative
    For i = 1 To Screen.FontCount
        Combo1.AddItem Screen.Fonts(i - 1)
        'Add the font list
    Next i
''''''''''''''''''''''''''''''''''''''''''''''''

    Combo1.text = Label1.FontName

End Sub
