VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGetPics 
   Caption         =   " Image Address Finder"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
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
   ScaleHeight     =   5205
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4950
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ComboBox txtURL 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   4875
   End
   Begin InetCtlsObjects.Inet inet1 
      Left            =   3720
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton command1 
      Caption         =   "&GetPics"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwOnlinePics 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5636
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Now click an image below to see if its the one that you want:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the website URL that has the image you want:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Right Justify
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   550
      Width           =   495
   End
End
Attribute VB_Name = "frmGetPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As String
Dim GetPicsfromSource As New CJString
Private Sub Command1_Click()
    On Error Resume Next
    Dim txt As String
    Dim b() As Byte
    Dim sourceLength As Integer
    Dim q As Integer

    Command1.Enabled = False
    
    StatusBar1.SimpleText = "Obtaining website source"
    
    b() = inet1.OpenURL(URL.Text, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    Source = txt
    
    Dim pics As String
    
    pics = ParsePageForPics(Source)
    
    lvwOnlinePics = pics
    
    Command1.Enabled = True
    Exit Sub
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

    lvwOnlinePics.ColumnHeaders(1).Width = (lvwOnlinePics.Width / 100) * 15
    lvwOnlinePics.ColumnHeaders(2).Width = (lvwOnlinePics.Width / 100) * 84

End Sub


Private Sub lvwOnlinePics_DblClick()
    
    Dim fldResult() As Byte
    
    strImageName = GetImageName(lvwOnlinePics.SelectedItem.SubItems(1))
    If strImageName <> "" Then
        fldResult = inetGetPics.OpenURL(lvwOnlinePics.SelectedItem.SubItems(1), icByteArray)
        If UBound(fldResult) > 0 Then
            Open App.Path & "\" & strImageName For Binary Access Write As #1
            Put #1, , fldResult()
            Close #1
            Load frmDisplay
            frmDisplay.Show
        End If
    End If
   
End Sub
