VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Image"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Do You want to use this image?"
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
      Left            =   360
      TabIndex        =   2
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Image imgPic 
      Height          =   4335
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "DoubleClick to close!!!"
      Top             =   0
      Width           =   4275
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    imgPic.Picture = LoadPicture(App.Path & "\" & strImageName)
    imgPic.Refresh
    
End Sub


Private Sub imgPic_DblClick()

    Kill App.Path & "\" & strImageName
    Unload frmDisplay

End Sub
