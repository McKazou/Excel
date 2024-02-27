Attribute VB_Name = "UpdateNameInSelection"
'Macro to name and update name in case the name change
'It will goes in the column AD and find the corresponding Format.
'This macro need "Microsft Spripting Runtime"
'Format accepted:
'|Cell 1    |Cell 2    |
'-----------------------
'|RUSK_01D01|        56|
'-----------------------
'|RUSK_01D02|      #NUM|
'-----------------------
'|RUSK_01S01|       C35|
'-----------------------
'Last changes :
'15/02/2019 : Added a condition for shared workbook; during sharing you can't edit name, you only create them so it will cause some duplicated if run on shared workbook
'15/02/2019 : Changed the default selected; was selecting the last cell instead of the Ad column
'26/02/2024 : Added a dictionnary to speed things up. Also solv why some were skipped
Private AllNamesDic


Sub UpdateNamesInSelection()
    Dim SelectedCell As Range
    Dim Celli As Range
    
    If Not ActiveWorkbook.MultiUserEditing Then
        Set SelectedCell = Selection
        Set SelectedCell = VerificationOnSelection(SelectedCell)
        
        If Not SelectedCell Is Nothing Then
        'getAllNames to avoid going trough each time
            Call getAllNames
    
            'Go through each cell and run the function to name a cell
            For Each Celli In SelectedCell
                If NameCellUsingLeftCell(Celli) Then
                    'MsgBox "Done for: " & Celli.Address
                End If
            Next Celli
            MsgBox "All Done"
        End If
    Else
        MsgBox "The current Excel is shared as Multi user Editing, you can't run this macro while the excel is shared"
    End If
    
    
    
End Sub
Function VerificationOnSelection(RangeSelected As Range) As Range

If TypeName(RangeSelected) = "Range" Then
            If RangeSelected.Areas.Count = 1 Then
                If RangeSelected.Columns.Count = 1 Then
                    If RangeSelected.Rows.Count > 1 Then
                        MsgBox "Selected Area used"
                        Set VerificationOnSelection = RangeSelected
                    Else
                        MsgBox "Default column selected (AD)"
                        LastLigne = Range("AD" & Rows.Count).End(xlUp).Row + 1
                        Set VerificationOnSelection = Range("AD1:AD" & LastLigne)
                    End If
                Else
                    MsgBox "Select only one column; LAZI"
                End If
            Else
                MsgBox "Select only one area; yeah sorry i was lazy to handle multiple areas"
            End If
        Else
        MsgBox "This selection isn't a range; what the fuck you think you are doing ?"
        End If
End Function


'This function will name the selected cell usign the text on the left.
Function NameCellUsingLeftCell(SelectedCell As Range) As Boolean

    Dim SelectedCellAddress As String
    Dim SelectedCellName As Name
    
            SelectedCell.Select 'TROUBLE SHOOTING
            If SelectedCellIsValid(SelectedCell) Then
                If SelectedCell.Offset(0, -1).Value <> 0 Then
                
                    'Trouble shooting
                    'Range("AE" & (105 + j)).Value = SelectedCell.Offset(0, -1).Value
                    'Range("AF" & (105 + j)).Value = SelectedCell.Parent.Name & "!" & SelectedCell.Address
                
                    'Recreate the address
                    SelectedCellAddress = "=" & SelectedCell.Parent.Name & "!" & SelectedCell.Address
                
                    Set SelectedCellName = isInNamesList(SelectedCellAddress)
                    If SelectedCellName Is Nothing Then
                        'Create Name
                        'Range("AG" & (105 + j)).Value = "Create Name" 'Trouble Shooting
                        ActiveWorkbook.Names.Add Name:=SelectedCell.Offset(0, -1).Value, RefersTo:=SelectedCell
                        'MsgBox "Name created: " & SelectedCell.Address & "using the name: " & SelectedCell.Offset(0, -1).Value
                    Else
                        'Range("AG" & (105 + j)).Value = "Modify" 'Trouble Shooting
                        SelectedCellName.Name = SelectedCell.Offset(0, -1).Value
                        'MsgBox "Name modified: " & SelectedCell.Address & " using the name: " & SelectedCell.Offset(0, -1).Value
                    End If
                    
                    'j = j + 1 'Trouble Shooting
                    NameCellUsingLeftCell = True
                End If
            End If
End Function

'Verification if the name is already exist or not
'If the name exist it will output the name so it can be modify.
Function isInNamesList(CellAddress As String) As Name
    'MsgBox CellAddress
    With ActiveSheet
        If AllNamesDic.Exists(CellAddress) Then
            Set isInNamesList = AllNamesDic(CellAddress)
        'For Each nName In ThisWorkbook.Names
        '    'MsgBox nName.RefersTo & " : " & CellAddress
        '    If Not (Contains = InStr(nName.RefersTo, .Name)) Then
        '        If Not (Contains = InStr(nName.RefersTo, CellAddress)) Then
        '            Set isInNamesList = nName
        '            Exit Function
        '        End If
        '    End If
        'Next
        End If
    End With
End Function

'Test if the cell is in error, if yes then you can name it
'       (it say the cell have a value but the value is currently bugged so you can name it but not refresh NX)
'If the cell isn't in error, test if it's a number and different of 0 then you can name it
'       (Different from zero in case of empty cells)
'If the cell isn't the condition above, test if it's a texte and different from empty then you can name it
Function SelectedCellIsValid(SelectedCell As Range) As Boolean
    'Cell is error ?
    If IsError(SelectedCell) Then
        SelectedCellIsValid = True
    'Cell Is a number ?
    ElseIf IsNumeric(SelectedCell.Value) Then
        'If the cell on the left is empty also then ignore it
        Dim offsetCell
        offsetCell = SelectedCell.Offset(0, -1).Value2
        If Not offsetCell = "" Then
            If Not IsNumeric(offsetCell) Then
                If Not offsetCell = "Nb of variables" Then
                    SelectedCellIsValid = True
                End If
            End If
        End If
    'Cell is a Text
    ElseIf Application.IsText(SelectedCell.Value) Then
            'Cell is a text not empty
            If SelectedCell.Value <> "" Then
            SelectedCellIsValid = True
            End If
    End If

End Function

'Function to get the intersection of two range
'return the range intersection or FALSE if not found
Private Sub getAllNames()
        'Store All names in dictionnary so we test if we find it while going trough cells we want to name
        Dim nm
        Dim wb As Workbook
        Set wb = ActiveWorkbook
        Dim nmRefersString As String
        Dim NamesDic                   'Create a variable
        Set NamesDic = CreateObject("Scripting.Dictionary")

        'Store all names
        For Each nm In wb.Names
            If nm.Visible Then
                nmRefersString = nm.RefersTo
                If Not NamesDic.Exists(nmRefersString) Then
                NamesDic.Add nmRefersString, nm
                Else
                    Debug.Print "[UpdateNamesInSelection] getAllNames - DuplicateNameFound" & nm.Name
                End If
            End If
        Next nm
        Set AllNamesDic = NamesDic
End Sub


