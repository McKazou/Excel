Attribute VB_Name = "CallSpreadsheetFunction"

Private SpreadsheetFunctionList As Object
Public DEFAULT_TABLE_CALCULATION_DATABASE_NAME As String

'This function will call a function by it's name,
'name found in the database in this workbook
'Database give where the file is
'Each time a file is opened it not closed right after
Function CallFunction(functionName As String, inputRange As Range) As Range
 On Error GoTo ErrorHandler
    Call Init
    Dim functionWB As Workbook
    If Not SpreadsheetFunctionList.Exists(functionName) Then
        'Open workbook store it in the dico
        Dim path As String
 
        path = getFilePathByName(functionName)
        Set functionWB = openWorkbook(path, True, False)
        SpreadsheetFunctionList.Add functionName, functionWB
    Else
        'Already open
        Set functionWB = SpreadsheetFunctionList(functionName)
    End If
 'we can put the input array in the inputs table
    'Application.Calculation = xlCalculationManual
    'Dim inputArray() As Variant
    'inputArray = inputRange.Value2
    'Call UpdateInputValues(functionWB, inputArray)
    
    'Application.Calculation = xlCalculationAutomatic
    
    'Dim outputTable As Range
    'outputTable = FindTable(functionWB, "Output")
    
    'CallFunction = outputTable.Value
    
    Exit Function
ErrorHandler:
        Application.Calculation = xlCalculationAutomatic
        'MsgBox "[CallFunction] Erreur lors de l'appel de la fonction : " & Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
 End Function




'Validé 240328
Private Sub Init()
    On Error GoTo ErrorHandler
        If SpreadsheetFunctionList Is Nothing Then
            Set SpreadsheetFunctionList = CreateObject("Scripting.Dictionary")
        End If
        DEFAULT_TABLE_CALCULATION_DATABASE_NAME = "Table_Functions_List"
    Exit Sub
ErrorHandler:
    'MsgBox "Erreur lors de l'initialisation du dictionnaire : " & Err.Description
End Sub

'Validé 240328
Private Function getFilePathByName(name As String)
  ' Obtenir le nom du classeur qui a appelé la fonction
    'Dim nomClasseur As String
    'nomClasseur = Application.Caller.Worksheet.Parent.name
    Call Init
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim folderPathCol As Range
    Dim nameCol As Range
    Dim i As Long
    Dim count As Long
    Dim tableFound As Boolean

    ' Récupérer le classeur qui a lancé la macro
    Set wb = ThisWorkbook

    ' Parcourir chaque feuille de calcul du classeur
    For Each ws In wb.Worksheets
        ' Parcourir chaque tableau de la feuille de calcul
        For Each tbl In ws.ListObjects
            ' Si le tableau est "Table_Fonctions_List"
            If tbl.name = DEFAULT_TABLE_CALCULATION_DATABASE_NAME Then
                ' Indiquer que le tableau a été trouvé
                tableFound = True
                ' Trouver la colonne "Folder Path"
                Set folderPathCol = tbl.ListColumns("Folder Path").Range
                ' Trouver la colonne "Name"
                Set nameCol = tbl.ListColumns("Name").Range
                
                ' Parcourir chaque ligne du tableau
                For i = 2 To nameCol.Rows.count
                    ' Si le nom correspond
                    If nameCol.Cells(i, 1).Value = name Then
                        ' Incrémenter le compteur
                        count = count + 1
                        ' Si le nom est trouvé pour la première fois
                        If count = 1 Then
                            ' Récupérer le "Folder Path" correspondant
                            getFilePathByName = folderPathCol.Cells(i, 1).Value & name
                        ' Si le nom est trouvé plus d'une fois
                        ElseIf count > 1 Then
                            ' Générer une erreur
                            Err.Raise 1004, , "Nom trouvé plusieurs fois"
                            Exit Function
                        End If
                    End If
                Next i
            End If
        Next tbl
    Next ws

    ' Si le tableau n'est pas trouvé, générer une erreur
    If Not tableFound Then
        Err.Raise 1004, , "[getFilePathByName] Tableau non trouvé"
    End If

    ' Si le nom n'est pas trouvé, générer une erreur
    If count = 0 Then
        Err.Raise 1004, , "[getFilePathByName] Nom non trouvé"
    End If
    
End Function
Sub opentest()
    Dim wb As Workbook
    Set wb = openWorkbook("\\stccwp0015\Worksresearsh$\01_Component Librairies\70_NdC et Tolérances\NX\getExportStringForNX\getExportStringForNX_240312.xlsx", _
        False, False)

End Sub

Public Function openFileFromCell(path As String)
    Dim wb As Workbook
    Set wb = openWorkbook(path)
    openFileFromCell = wb.name
End Function
'Validé 240328
Private Function openWorkbook(path As String, Optional readOnly As Boolean = False, Optional hidden As Boolean = True) As Workbook
    Dim methodToOpen As String
    'methodToOpen = "Application.Workbooks.Open"
    methodToOpen = "Shell"
    Dim wb As Workbook
    
    If path <> "" Then
        If Dir(path) <> "" Then ' Check if file exists
            Set wb = GetOpenWorkbook(path)
            If wb Is Nothing Then ' Check if workbook is already open
                Application.EnableEvents = False ' Disable events
                Select Case methodToOpen
                    Case "Application.Workbooks.Open"
                    'Doesn't fucking work when call from a cell
                        Set wb = Application.Workbooks.Open(path, , readOnly)
                        'Set openWorkbook = Application.Workbooks.Open(path, , readOnly)

                    Case "Shell"
                    'This work from a celle but now i need to "find the instance"
                        Dim cmd As String
                        'Sans nouvelle instance
                        cmd = "cmd /c start """" ""excel.exe"" """ & path & """"
                        'avec nouvelle instance
                        'cmd = "cmd /c start """" ""excel.exe /x"" """ & path & """"
                        Call Shell(cmd, vbNormalFocus)
                        
                         Application.Wait (Now + TimeValue("0:00:10"))
                        ' Wait for the workbook to open
                        Set wb = waitForExcelToOpen(path)
                        Application.EnableEvents = True ' Enable events
                    Case Else
                End Select
            End If
            
            'MsgBox "excel opened"
                
            If wb Is Nothing Then
                Err.Raise 1005, , "[openWorkbook] Error: The Object Workbook is nothing with the path: " & path
            Else
                Dim window As Variant
                For Each window In wb.Windows
                    window.Visible = Not hidden
                Next window
                Set openWorkbook = wb
            End If
        Else
            Err.Raise 1007, , "[openWorkbook] Error: The file does not exist at the specified path."
        End If
    Else
        Err.Raise 1008, , "[openWorkbook] Error: No file path provided."
    End If
End Function


'This will way max 10 sec for an excel to open
Private Function waitForExcelToOpen(workbookName As String) As Workbook
    Dim start As Double
    Dim timeout As Double
    Dim elapsedTime As Double
    
    start = Timer
    timeout = 10 ' Set a timeout (in seconds) to avoid an infinite loop
    
    Do
        elapsedTime = Timer - start
        If elapsedTime > timeout Then Exit Do
        On Error Resume Next
        Set waitForExcelToOpen = GetObject(workbookName)
        On Error GoTo 0
        If Not waitForExcelToOpen Is Nothing Then Exit Do
        DoEvents
    Loop
End Function


' Function to check if a workbook is open
Function GetOpenWorkbook(pathOrFileName As String) As Workbook
    Dim FileName As String
    FileName = Dir(pathOrFileName)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name = FileName Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOpenWorkbook = Nothing
End Function



Sub cleanExcelLoaded()
 On Error GoTo ErrorHandler
    Dim win As window
    For Each win In Application.Windows
        If win.Caption = ThisWorkbook.name Then
            ' Ne ferme pas le classeur qui a lancé la fonction
        Else
            win.Activate
            ActiveWindow.Close saveChanges:=saveModification
        End If
    Next
    Exit Sub
ErrorHandler:
    'MsgBox "Erreur lors du cleanup des excel : " & Err.Description
End Sub



'This will find the table in a workbook and return the rnage associated
Private Function FindTable(wb As Workbook, tableName As String) As Range
    Dim ws As Worksheet
    Dim tbl As ListObject

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.name = tableName Then
                Set FindTable = tbl.Range
                Exit Function
            End If
        Next tbl
    Next ws
    
    Err.Raise Number:=vbObjectError + 4, _
                Description:="[FindTable] cannot find the table named:" & tableName
    Set FindTable = Nothing
End Function

Sub UpdateInputValues(wb As Workbook, inputArray() As Variant)
    Dim rng As Range
    Dim i As Long, j As Long
    Dim matchCount As Long

    Set rng = FindTable(wb, "Input")

    For i = LBound(inputArray, 1) To UBound(inputArray, 1)
        matchCount = 0
        For j = 2 To rng.Rows.count ' Commence à 2 pour ignorer l'en-tête
            If rng.Cells(j, 1).Value = inputArray(i, 1) Then
                rng.Cells(j, 2).Value = inputArray(i, 2)
                matchCount = matchCount + 1
            End If
        Next j
        If matchCount = 0 Then
            Err.Raise Number:=vbObjectError + 9999, _
                       Description:="Erreur: Le paramètre " & inputArray(i, 1) & " n'a pas été trouvé dans Input_Name."
            Exit Sub
        ElseIf matchCount > 1 Then
            Err.Raise Number:=vbObjectError + 9999, _
                       Description:="Erreur: Le paramètre " & inputArray(i, 1) & " est présent plus d'une fois dans Input_Name."
            Exit Sub
        End If
    Next i
End Sub

