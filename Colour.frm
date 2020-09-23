VERSION 5.00
Begin VB.Form ColorForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Select Colour"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7560
   Begin VB.CommandButton Command2 
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
      Left            =   6120
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Colour"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "#000000"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Misc"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         Begin VB.CommandButton Command3 
            Caption         =   "Random"
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
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Black"
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
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "White"
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
            TabIndex        =   11
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   2295
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            Height          =   1695
            Left            =   120
            ScaleHeight     =   1635
            ScaleWidth      =   1995
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Red, Green, Blue"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   9
            Top             =   1440
            Width           =   255
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   8
            Top             =   840
            Width           =   255
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            LargeChange     =   10
            Left            =   360
            Max             =   255
            SmallChange     =   10
            TabIndex        =   4
            Top             =   1440
            Width           =   2895
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   10
            Left            =   360
            Max             =   255
            SmallChange     =   10
            TabIndex        =   3
            Top             =   840
            Width           =   2895
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   10
            Left            =   360
            Max             =   255
            SmallChange     =   10
            TabIndex        =   2
            Top             =   240
            Width           =   2895
         End
      End
   End
End
Attribute VB_Name = "ColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'''Pad the colour values so we get the right hex values
Private Function LeftPad(Value, Size As Long, Optional PadCharacter As String = " ") As String
    LeftPad = "" & Value
    While Len(LeftPad) < Size
        LeftPad = PadCharacter & LeftPad
    Wend
End Function

Private Sub command1_Click()
HexCode = Text8.text 'load the colour into the string
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

'''''''''''''''''''''''
Randomize
Random = Int(Rnd * 255)
HScroll1.Value = Random
Randomize
Random = Int(Rnd * 255)
HScroll2.Value = Random
Randomize
Random = Int(Rnd * 255)
HScroll3.Value = Random
'This just randomizes our 3 colours to get a random colour
'''''''''''''''''''''''
End Sub

Private Sub Command6_Click()
'Set colours to 0, thus giving us the colour black
HScroll1.Value = 0
HScroll2.Value = 0
HScroll3.Value = 0
End Sub

Private Sub Command7_Click()
'Set all colours to maximum to give us white
HScroll1.Value = 255
HScroll2.Value = 255
HScroll3.Value = 255
End Sub

Private Sub HScroll1_Change()
''Change the back colour to the colour we have selected. We update this as we scroll (real time)
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

'''Turn our colour into Hex, and display it in the text box
Text8.text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change 'When we scroll, we call Hscroll1_change, so it updates the colour
End Sub

Private Sub HScroll2_Change()
''Change the back colour to the colour we have selected. We update this as we scroll (real time)
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

'''Turn our colour into Hex, and display it in the text box
Text8.text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change 'When we scroll, we call Hscroll1_change, so it updates the colour
End Sub

Private Sub HScroll3_Change()
''Change the back colour to the colour we have selected. We update this as we scroll (real time)
Picture1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)

'''Turn our colour into Hex, and display it in the text box
Text8.text = "#" & LeftPad(Hex(HScroll1.Value), 2, 0) & LeftPad(Hex(HScroll2.Value), 2, 0) & LeftPad(Hex(HScroll3.Value), 2, 0)
End Sub

Private Sub HScroll3_Scroll()
HScroll3_Change 'When we scroll, we call Hscroll1_change, so it updates the colour
End Sub

