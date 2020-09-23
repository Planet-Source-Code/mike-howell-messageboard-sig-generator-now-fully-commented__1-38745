VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'********************************************************************************************
'Functions:
'********************************************************************************************

'Local variable(s)to hold property value(s)
Private mVarWorkingString   As String           'OBJ.WorkingString
Private mVarStringLength    As Long             'OBJ.StringLength

Enum PreserveVal
     DoPreserve = 0
     DontPreserve = 1
End Enum
    

Public Function TrimStringRight() As String
    
    mVarWorkingString = RTrim$(mVarWorkingString)
    TrimStringRight = mVarWorkingString
    
End Function

Public Function TrimStringBoth() As String
    
    mVarWorkingString = LTrim$(mVarWorkingString)
    mVarWorkingString = RTrim$(mVarWorkingString)
    mVarWorkingString = Trim$(mVarWorkingString)
    
    TrimStringBoth = mVarWorkingString
    
End Function


Public Function TrimStringLeft() As String
    
    mVarWorkingString = LTrim$(mVarWorkingString)
    TrimStringLeft = mVarWorkingString
    
End Function

Public Property Let WorkingString(ByVal vData As String)

    mVarWorkingString = vData
    Length
    
End Property


Public Property Get StringLength() As String

    StringLength = mVarStringLength
    
End Property

Public Property Get WorkingString() As String

    WorkingString = mVarWorkingString
    
End Property


Private Sub Class_Initialize()
    
    mVarWorkingString = ""
    mVarStringLength = 0
    
End Sub



Public Function FindAndReplace(szFilter As String, szReplace As String)
    
    mVarWorkingString = Replace(mVarWorkingString, szFilter, szReplace)
    
End Function

Public Function Length() As Long

    mVarStringLength = Len(mVarWorkingString)
    Length = mVarStringLength
    
End Function

Private Sub Class_Terminate()
    
    mVarWorkingString = ""
    
End Sub



Public Function CheckInvalidCharacters(chrArray, iStartPos As Long)

Dim iArrCounter         As Integer
Dim szCompareString     As String
Dim lnFoundPos           As Long

If IsArray(chrArray) Then

    For iArrCounter = 0 To UBound(chrArray)
    
        If chrArray(iArrCounter) = "" Then
            CheckInvalidCharacters = 0
            Exit Function
        Else
            szCompareString = CStr(chrArray(iArrCounter))
            
            lnFoundPos = InStr(iStartPos, mVarWorkingString, szCompareString, 1)
            If lnFoundPos <> 0 Then
                'Invalid character has been found
                CheckInvalidCharacters = lnFoundPos
                Exit Function
            Else
                'No invalid character found, continue
                CheckInvalidCharacters = lnFoundPos
            End If
        End If
        
        szCompareString = ""
        lnFoundPos = 0
        
    Next iArrCounter

Else
    
    CheckInvalidCharacters = 0

End If

End Function

Public Function Switch2Letters(lPos1 As Long, lPos2 As Long)

Dim szChr1                  As String
Dim szChr2                  As String
Dim szCat1                  As String
Dim lLen                    As Long
Dim lnCount                 As Long
Dim szChrWorkingString      As String
Dim strWorkingString        As String

lLen = Length
If lPos1 <= 0 Or lPos1 > lLen Or lPos2 <= 0 Or lPos2 > lLen Then
    Exit Function
End If

szChr1 = Mid(mVarWorkingString, lPos1, 1)
szChr2 = Mid(mVarWorkingString, lPos2, 1)


For lnCount = 1 To lLen
    
    szChrWorkingString = Mid(mVarWorkingString, lnCount, 1)
    
    If lnCount = lPos1 Then
        strWorkingString = strWorkingString & szChr2
        
    ElseIf lnCount = lPos2 Then
        strWorkingString = strWorkingString & szChr1
        
    Else
        strWorkingString = strWorkingString & szChrWorkingString
    End If

Next lnCount
    
mVarWorkingString = strWorkingString

End Function

Public Function ReverseLetters()

If mVarStringLength = 0 Then
    Exit Function
Else

    Dim lnX As Long
    Dim szChr As String
    Dim MyStr As String
    
    Do
        szChr = Mid(mVarWorkingString, mVarStringLength - lnX, 1)
        MyStr = MyStr & szChr
        
        lnX = lnX + 1
    Loop Until lnX = mVarStringLength

    mVarWorkingString = MyStr
    
End If

End Function


Public Function MakeArrayOfChrs() As String()

    Dim arrayOfChrs()           As String
    Dim x                       As Integer
    Dim myChr                   As String
    
    ReDim arrayOfChrs(Len(mVarWorkingString))
    x = 1
    Do
        arrayOfChrs(x - 1) = Mid(mVarWorkingString, x, 1)
        x = x + 1
    
    Loop Until x - 1 = Len(mVarWorkingString)
    
    ReDim Preserve arrayOfChrs(UBound(arrayOfChrs) - 1)
    MakeArrayOfChrs = arrayOfChrs()
    
End Function
