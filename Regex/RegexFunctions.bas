Attribute VB_Name = "RegexFunctions"
Option Explicit

'this regex function will replace all string in another using a pattern. If the pattern isn't found the string will stay the same
Function RegexReplace(StrInput As String, strPattern As String, replacedBy As String) As String
    Dim Regex As New RegExp
    Dim strReplace As String
    Dim strOutput As String

    If strPattern <> "" Then
        
        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(StrInput) Then
            RegexReplace = Regex.Replace(StrInput, replacedBy)
        Else
            RegexReplace = StrInput
        End If
    End If
    Set Regex = Nothing
End Function

Function RegexReplaceFirst(StrInput As String, strPattern As String, replacedBy As String) As String
    Dim Regex As New RegExp
    Dim strReplace As String
    Dim strOutput As String

    If strPattern <> "" Then
        
        With Regex
            .Global = False
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(StrInput) Then
            RegexReplaceFirst = Regex.Replace(StrInput, replacedBy)
        Else
            RegexReplaceFirst = StrInput
        End If
    End If
    Set Regex = Nothing
End Function

Function StringReverse(s As String)
    StringReverse = StrReverse(s)
End Function

Function RegexFindFirst(StrInput As String, strPattern As String) As String
On Error GoTo errHandler
    Dim Regex As New RegExp
    Dim strReplace As String
    Dim strOutput As String
    
    If strPattern <> "" Then
        
        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        If Regex.Test(StrInput) Then
            RegexFindFirst = Regex.Execute(StrInput)(0)
            
        Else
            'return the same string
            RegexFindFirst = "#NOTFOUND"
            'return an error
        End If
    End If
    Exit Function
    
errHandler:
    Select Case Err.Number
        Case 5017
        RegexFindFirst = "#PATTERN"
        Case Else
            MsgBox "There's some problem with the value you have in the cell A1." & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Description: " & Err.Description
    End Select
End Function


Function RegexRangeFindFirst(RangeInput As Range, strPattern As String) As Variant()
'On Error GoTo errHandler

    'https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba
    Dim InputArray As Variant
    InputArray = RangeInput.Value

    Dim resultRange() As Variant
    ReDim resultRange(UBound(InputArray) - 1)
    
    Dim i As Integer
    i = 0
    
    If strPattern <> "" Then
        Dim cel As Range
        For Each cel In RangeInput
                resultRange(i) = RegexFindFirst(cel.Value, strPattern)
                i = i + 1
        Next
        'https://stackoverflow.com/questions/24456328/creating-and-transposing-array-in-vba
        RegexRangeFindFirst = Application.Transpose(resultRange)
    End If
    Exit Function
End Function

Function RegexTest(StrInput As String, strPattern As String)
On Error GoTo errHandler
    Dim Regex As New RegExp
    Dim strReplace As String
    Dim strOutput As String
    
    If strPattern <> "" Then
        
        With Regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        RegexTest = Regex.Test(StrInput)

    End If
    Exit Function
    
errHandler:
    Select Case Err.Number
        Case 5017
        RegexTest = "#PATTERN"
        Case Else
            MsgBox "There's some problem with the value you have in the cell A1." & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Description: " & Err.Description
    End Select
End Function

