VERSION 5.00
Begin VB.Form ImportantEzInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Important Information"
   ClientHeight    =   1680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "ImportantEzInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Do not show this message again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.CommandButton command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"ImportantEzInfo.frx":0442
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "ImportantEzInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub

Private Sub command1_Click()
'if check1 is ticked, then write to the registry
'saying that the user doesent want to see the dialog
'again
If Check1.Value = 1 Then
SetStringValue "HKEY_CURRENT_USER\Software\Howelly\MSG", "ShowEzInfoDialog", "No"
End If
''''''''''''''''''''''''''''''

Unload Me
End Sub
