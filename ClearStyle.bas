Attribute VB_Name = "ClearStyle"
Dim StyleCount As Integer 'Number of style

'Objectif: Clear all duplicated style in the style list will delete anything that god a "Space"&10 or 1&"Space"&1 at the end
'Upgrade trajectory : Merge to the one without
Sub ClearStyle()
    CptDeleted = 0
    i = 1
    
    For Each IStyle In ActiveWorkbook.Styles
        StyleCount = ActiveWorkbook.Styles.Count
        Application.StatusBar = "Style count : " & StyleCount & " / " & i & " Deleted: " & CptDeleted
        If ToKeep(IStyle.Name) Then
            'Worksheets("Sheet1").Cells(i, 1) = IStyle.Name
        Else
            IStyle.Delete
            CptDeleted = CptDeleted + 1
        End If
        i = i + 1
    Next
End Sub

Function ToKeep(StyleName As String) As Boolean
    ToKeep = True
    RightString = Right(StyleName, 3)
    FoundSpace = InStr(1, RightString, " ")
    If FoundSpace >= 1 Then
        ToKeep = False
    End If
End Function
