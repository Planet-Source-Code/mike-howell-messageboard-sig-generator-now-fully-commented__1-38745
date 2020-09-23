VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BackColor       =   &H005A595A&
   Caption         =   " Messageboard Sig Generator"
   ClientHeight    =   7605
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   11880
   ClipControls    =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEndHTML 
      Height          =   3285
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Main.frx":0ABA
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtStartHTML 
      Height          =   3285
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Main.frx":24E4
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1535
      ButtonWidth     =   609
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Text            =   "Unsaved"
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Extra Code"
         Height          =   840
         Left            =   11040
         Picture         =   "Main.frx":5888
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Extra HTML & Java code"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Advanced"
         Height          =   840
         Left            =   10080
         Picture         =   "Main.frx":5CCA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Advance Tag Information"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Add List"
         Height          =   840
         Left            =   9120
         Picture         =   "Main.frx":6674
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Add in a list"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Table"
         Height          =   840
         Left            =   8160
         Picture         =   "Main.frx":76B6
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Add a table"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Video"
         Height          =   840
         Left            =   7200
         Picture         =   "Main.frx":7AF8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add a video"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add Sound File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6240
         Picture         =   "Main.frx":7E02
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add a sound file"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Horizontal line"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5280
         Picture         =   "Main.frx":7F4C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Add a horizontal line"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Next Line <br>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4440
         Picture         =   "Main.frx":82D6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add a linebreak"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add Link"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3600
         Picture         =   "Main.frx":8420
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add a Hyperlink"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add Image"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2640
         Picture         =   "Main.frx":872A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add an image"
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Text"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1800
         Picture         =   "Main.frx":8A34
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add Text"
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton Command17 
         Height          =   420
         Left            =   1440
         Picture         =   "Main.frx":8B7E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Height          =   420
         Left            =   1080
         Picture         =   "Main.frx":8CC8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Paste"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Height          =   420
         Left            =   720
         Picture         =   "Main.frx":8E12
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Copy"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Height          =   420
         Left            =   360
         Picture         =   "Main.frx":8F5C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cut"
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command16 
         Height          =   420
         Left            =   0
         Picture         =   "Main.frx":90A6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "New"
         Top             =   420
         Width           =   375
      End
      Begin VB.Line Line6 
         BorderWidth     =   4
         X1              =   10080
         X2              =   10080
         Y1              =   0
         Y2              =   840
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Sig"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E3CCBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00904D36&
      Height          =   5415
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Main.frx":91F0
      Top             =   1320
      Width           =   10935
   End
   Begin VB.Line Line5 
      BorderWidth     =   5
      X1              =   0
      X2              =   10800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   480
      X2              =   11400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   11400
      X2              =   11400
      Y1              =   1320
      Y2              =   6720
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   480
      X2              =   11400
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   480
      X2              =   480
      Y1              =   1320
      Y2              =   6720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveasWP 
         Caption         =   "Save As Webpage"
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnutest 
         Caption         =   "Test Sig"
      End
      Begin VB.Menu mnuPrintCode 
         Caption         =   "Print Sig Code"
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSA 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Below is simply declaring the api and declaring the
'constants, you will either understand this or you
'wont :)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Const WS_CHILD = &H40000000
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const SW_HIDE = 0
Const SW_NORMAL = 1
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim MD As String
Dim tWnd As Long, bWnd As Long, ncWnd As Long

Private Sub command1_Click()
Dim strshowMes As String

'This is to see if the user has seen this
'message before, if the user has seen it and choosen
'not to see it again, then it saves a setting in
'the registry.

strshowMes = GetStringValue("HKEY_CURRENT_USER\Software\Howelly\MSG", "ShowEzInfoDialog") ' This code reads the registry to see if the user wants to see the message again


If strshowMes = "Yes" Then ' if the user wants to see the message again then:
ImportantEzInfo.Show 1
'show the form with the message in.
'if you put '1' after the '.show' it makes the form model
'making the form model, means no more code is executed, untill the
'form is closed. bassically like a messagebox
End If


Dim test As String
Dim temppath As String
Dim IE As Long

test = txtStartHTML.text & Text1.text & txtEndHTML.text ' This load all the HTML into a string we need to test the sig.

temppath = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") & "\MSGTemp" 'This loads a temperoray path into the string. The temporary path is the My Documents path (extracted from the registry) plus the programs folder for temporary files, which is deleted when the program is closed

Open temppath & "\MsgBoardSigTest(" & TestTemp & ").html" For Append As 1 'This created a blank file, in our temporary folder
Print #1, test 'writes the HTML code we loaded into the string earlier, into the file
Close 1 'close the file, ready for use

IE& = ShellExecute(Me.hwnd, "open", temppath & "\MsgBoardSigTest(" & TestTemp & ").html", vbNullString, vbNullString, SW_MAXIMIZE) 'This opens our file we just saved into an IE window

TestTemp = TestTemp + 1 'Add one to the testtemp so that we dont overright files

End Sub

Private Sub Command10_Click()
Text1.SetFocus
Main.Hide
CodeWindow.Show
End Sub

Private Sub Command11_Click()
Text1.SetFocus
Main.Hide
Advanced.Show

End Sub

Private Sub Command12_Click()
ListForm.Show 1
Text1.text = Text1.text & ListCode
ListCode = ""
Text1.SetFocus
End Sub

Private Sub Command13_Click()
Clipboard.SetText Text1.text
Text1.SetFocus
End Sub

Private Sub Command14_Click()
Clipboard.SetText Text1.text
Text1.text = ""
Text1.SetFocus
End Sub

Private Sub Command15_Click()
Text1.text = Text1.text & Clipboard.GetText
Text1.SetFocus
End Sub

Private Sub Command16_Click()
Dim Y As String

If Text1.text = "<br>" Then 'if text1 still has only <br>
Exit Sub 'exit the sub because nothing has been entered from the programs original state
End If



Y = MsgBox("Do you want to clear the current HTML?", vbYesNo, "New") 'Check if the user really wants to clear the HTML

If Y <> vbYes Then 'Checks the users answer, and carried out the code accordingly
Exit Sub
End If

Text1.text = "<br>" 'set the program back to its original state
Text1.SetFocus
End Sub

Private Sub Command17_Click()
Call PrintText(Text1) 'Prints the text to the printer
Text1.SetFocus
End Sub

Private Sub Command19_Click()
Text1.Refresh
End Sub

Private Sub Command2_Click()
ImageForm.Show 1
Text1.text = Text1.text & ImageHTMLCode 'update the main HTML code with the code we just obtanied from the user
ImageHTMLCode = ""
Text1.SetFocus
End Sub

Private Sub Command20_Click()

End Sub

Private Sub Command21_Click()


End Sub

Private Sub Command22_Click()
FontLists.Show 1
Text1.FontName = strFontName
strFontName = ""
End Sub

Private Sub Command3_Click()
Text1.text = Text1.text & "<br>" & vbCrLf
Text1.SetFocus
End Sub

Private Sub Command4_Click()
Linkform.Show 1
Text1.text = Text1.text & LinkHTMLCode 'update the main HTML code with the code we just obtanied from the user
LinkHTMLCode = ""
Text1.SetFocus
End Sub

Private Sub Command5_Click()
textform.Show 1

Text1.text = Text1.text & FontHTMLCode 'update the main HTML code with the code we just obtanied from the user
FontHTMLCode = ""
Text1.SetFocus
End Sub

Private Sub Command6_Click()
hrline.Show 1

Text1.text = Text1.text & HRLineCode 'update the main HTML code with the code we just obtanied from the user

HRLineCode = ""

Text1.SetFocus
End Sub

Private Sub Command7_Click()
BGSound.Show 1
Text1.text = Text1.text & BGSoundCode 'update the main HTML code with the code we just obtanied from the user
BGSoundCode = ""
Text1.SetFocus
End Sub

Private Sub Command8_Click()
VideoForm.Show 1
Text1.text = Text1.text & VideoCode 'update the main HTML code with the code we just obtanied from the user
VideoCode = ""
Text1.SetFocus
End Sub



Private Sub Command9_Click()
TableForm.Show 1
Text1.text = Text1.text & TableCode 'update the main HTML code with the code we just obtanied from the user
TableCode = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next


    ' If we're being debugged, then stop execution!
    If IsDebuggerPresent <> 0 Then
    MsgBox "No debugging this program", vbExclamation, "Error"
    End
    End If


Text1.text = "<br>" & vbCrLf 'set up the text box
TestTemp = 1 ' set the number of our temporary files
MD = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal") & "\MSGTemp" ' Load the path for our temporary folder into a string
MkDir MD 'Make the temporary folder
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill MD & "\*.*" 'Delete the contents of our temporary folder
RemoveDirectory MD 'Delete the temporary folder
End Sub
Private Sub mnuabout_Click()
MsgBox "This program is copyright of Michael Howell", vbInformation, "Copyright"
End Sub

Private Sub mnuClear_Click()

End Sub

Private Sub mnuclose_Click()
End
End Sub

Private Sub mnuCopy_Click()
Call Command13_Click
End Sub

Private Sub mnuCut_Click()
Call Command14_Click
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnunew_Click()
Dim val As String

val = MsgBox("Are you sure you want to clear all the HTML?", vbYesNo, "New")

If val <> 6 Then
Exit Sub
End If

Text1.text = "<br>"

End Sub



Private Sub mnuOpen_Click()
Dim q As String

'Validates if the user wants to load a sig
q = MsgBox("Do you want to load a sig?", vbYesNo, "Save")
If q <> vbYes Then
Exit Sub
End If

CommonDialog1.Filter = "Text Files (*.txt)|*.txt|MSG Files (*.msg)|*.mesg|" 'Set the filter of our common dialog, so only text files, and MSG files are seen
CommonDialog1.ShowOpen ' open the common dialog

Call LoadText(Text1, CommonDialog1.FileName)


Combo1.AddItem CommonDialog1.FileTitle 'Add the filename to our combo box
Combo1.text = CommonDialog1.FileTitle
End Sub

Private Sub mnuPaste_Click()
Call Command15_Click
End Sub

Private Sub mnuPrintCode_Click()
Printer.Print Text1.text
Printer.EndDoc
End Sub

Private Sub mnuSave_Click()
Dim q As String

q = MsgBox("Do you want to save your current sig?", vbYesNo, "Save")
If q <> vbYes Then
Exit Sub
End If

    'Save Files Using CommonDialog Controls
    '--------------------------------------
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|ESG Files (*.esg)|*.esg|"
    CommonDialog1.ShowSave
    'If User has Selected A Valid File Then


    If CommonDialog1.FileName <> "" Then
        'Open The File For Output
Open CommonDialog1.FileName For Append As 1
Print #1, Text1.text
Close 1

    Else
    Exit Sub
    End If

Combo1.AddItem CommonDialog1.FileTitle
Combo1.text = CommonDialog1.FileTitle
End Sub

Private Sub mnuSaveasWP_Click()

'Save the sig, works under the samre prinicple as loading
Dim test As String
Dim temppath As String
Dim IE As Long

test = txtStartHTML.text & Text1.text & txtEndHTML.text

CommonDialog1.Filter = "HTML File (*.html)|*.html|"


CommonDialog1.ShowSave

temppath = CommonDialog1.FileName

Open temppath For Append As 1
Print #1, test
Close 1
End Sub

Private Sub mnutest_Click()
Call command1_Click
End Sub

Private Sub Timer1_Timer()


If Main.WindowState = 0 Then
Main.WindowState = 2
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Exit Sub
End Sub

Public Sub PrintText(text As TextBox)
    'Routine for printing, Very short, but very good.
    'Feel free to use
    
    Printer.Print "" + text.text + Str(Printer.Page)
    Printer.NewPage
    Printer.Print "" + text.text + Str(Printer.Page)
    Printer.EndDoc
End Sub
Sub LoadText(Destination As TextBox, FilePath As String)

    On Error GoTo error
    Dim MyStr As String
    Open file For Input As #1


    Do While Not EOF(1)
        Line Input #1, a$
        texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
    Loop
    Lst = texto$
    Close #1
    Exit Sub
error:
    X = MsgBox("File Not Found", vbOKOnly, "Error")
End Sub

