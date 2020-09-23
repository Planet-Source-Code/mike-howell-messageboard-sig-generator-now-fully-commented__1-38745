Attribute VB_Name = "modChooseColor"
Option Explicit

Private Declare Function ShowColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80

Public Function ColourSelect(hwndForm As Long, DefaultColour As Long) As Long
    
    Dim CC As CHOOSECOLOR
    Static strColour As String, i As Integer
    
    'lpCustColors requires a String representing RGB with a Terminator.
    'The string must be made up of the characters 0 to 255
    'to fill the first box with Yellow
    '           Red        Green    Blue     Terminator
    strColour = Chr(255) & Chr(255) & Chr(0) & Chr(0)
    'or alternatively
    'strColour = "ÿÿ"
    
    'there are 16 custom color boxes so the string should be 64 characters long.
    'You may make the string shorter, but as in the examples above you will get
    'random colors displayed in the other custom color boxes, unless you declare the
    'string as String * 64 and leave it empty, whereby you will fill all boxes black.
    
    'To fill all the boxes with white just pass this string
    'strColour = String(64, 255)
    
    With CC
        .lStructSize = Len(CC)
        .hWndOwner = hwndForm
        .rgbResult = DefaultColour
        .lpCustColors = strColour
        .flags = CC_FULLOPEN Or CC_RGBINIT Or CC_SOLIDCOLOR
    End With
    
    If ShowColor(CC) <> 0 Then
        ColourSelect = CC.rgbResult
    Else
        ColourSelect = DefaultColour
    End If
    'in order to remember your choice of custom colours you must save
    'strcolour here. If kept as is strColour will be overwritten the next time
    'this function is called.
    strColour = CC.lpCustColors
    
End Function

