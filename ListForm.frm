VERSION 5.00
Begin VB.Form ListForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " List"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
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
   ScaleHeight     =   3705
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "List Contents"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
      Begin VB.CheckBox Check1 
         Caption         =   "Yes"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Do you want this program to fill in the list for the contents you provide?"
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Enter how many many list items you would like:"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Type"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.OptionButton Option2 
         Caption         =   "Numbered List"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bulleted List"
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

'First we see what type of list a user wants. Numbers
'or bullet points
If Option1.Value = True Then
ListCode = "<ul>" & vbCrLf
End If

If Option2.Value = True Then
ListCode = "<ol>" & vbCrLf
End If
''''''''''''''''''''''''''''

'if the user selects to add his own content, then we
'just write the HTML code into the string, depending
'on how many items in the list the user wants.

If Check1.Value = 0 Then

    Dim i As Integer
    Dim total As Integer

    total = Text1.text

    For i = 1 To total
    ListCode = ListCode & "<li> <!-- Add Content Here --> </li>" & vbCrLf
    Next i
End If

'if the user asks us to add the content, then for each
'List item we show an inputbox and ask the user
'for what data he wants. When he selects what he wants
'we write it to the HTML code.
If Check1.Value = 1 Then

    Dim j As Integer
    Dim ListTotal As Integer

    ListTotal = Text1.text

    For j = 1 To ListTotal
    ListCode = ListCode & "<li>"
    ListCode = ListCode & InputBox("Enter your " & j & ". contents", "Content")
    ListCode = ListCode & "</li>" & vbCrLf
    Next j
End If
''''''''''''''''''''''''''''''

''Close the list tags accordingly.
If Option1.Value = True Then
ListCode = ListCode & "</ul>"
End If

If Option2.Value = True Then
ListCode = ListCode & "</ol>"
End If
''''

Unload Me
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
 '''Make sure no numbers can be entered into text1
    Dim Numbers As Integer
    Dim Msg As String
    Numbers = KeyAscii


    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        KeyAscii = 0
    End If
End Sub
