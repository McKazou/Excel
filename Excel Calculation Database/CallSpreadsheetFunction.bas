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
    End If
 'we can put the input array in the inputs table
    Application.Calculation = xlCalculationManual
    Dim inputArray() As Variant
    inputArray = inputRange.Value2
    Call UpdateInputValues(functionWB, inputArray)
    
    Application.Calculation = xlCalculationAutomatic
    
    Dim outputTable As Range
    outputTable = FindTable(functionWB, "Output")
    
    CallFunction = outputTable.Value
    
    Exit Function
    ErrorHandler:
        Application.Calculation = xlCalculationAutomatic
        MsgBox "[CallFunction] Erreur lors de l'appel de la fonction : " & Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
 End Function

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
        For j = 2 To rng.Rows.count ' Commence � 2 pour ignorer l'en-t�te
            If rng.Cells(j, 1).Value = inputArray(i, 1) Then
                rng.Cells(j, 2).Value = inputArray(i, 2)
                matchCount = matchCount + 1
            End If
        Next j
        If matchCount = 0 Then
            Err.Raise Number:=vbObjectError + 9999, _
                       Description:="Erreur: Le param�tre " & inputArray(i, 1) & " n'a pas �t� trouv� dans Input_Name."
            Exit Sub
        ElseIf matchCount > 1 Then
            Err.Raise Number:=vbObjectError + 9999, _
                       Description:="Erreur: Le param�tre " & inputArray(i, 1) & " est pr�sent plus d'une fois dans Input_Name."
            Exit Sub
        End If
    Next i
End Sub



'Valid� 240328
Private Sub Init()
    On Error GoTo ErrorHandler
        If SpreadsheetFunctionList Is Nothing Then
            Set SpreadsheetFunctionList = CreateObject("Scripting.Dictionary")
        End If
        DEFAULT_TABLE_CALCULATION_DATABASE_NAME = "Table_Functions_List"
    Exit Sub
    ErrorHandler:
    MsgBox "Erreur lors de l'initialisation du dictionnaire : " & Err.Description
End Sub

'Valid� 240328
Private Function getFilePathByName(name As String)
  ' Obtenir le nom du classeur qui a appel� la fonction
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

    ' R�cup�rer le classeur qui a lanc� la macro
    Set wb = ThisWorkbook

    ' Parcourir chaque feuille de calcul du classeur
    For Each ws In wb.Worksheets
        ' Parcourir chaque tableau de la feuille de calcul
        For Each tbl In ws.ListObjects
            ' Si le tableau est "Table_Fonctions_List"
            If tbl.name = DEFAULT_TABLE_CALCULATION_DATABASE_NAME Then
                ' Indiquer que le tableau a �t� trouv�
                tableFound = True
                ' Trouver la colonne "Folder Path"
                Set folderPathCol = tbl.ListColumns("Folder Path").Range
                ' Trouver la colonne "Name"
                Set nameCol = tbl.ListColumns("Name").Range
                
                ' Parcourir chaque ligne du tableau
                For i = 2 To nameCol.Rows.count
                    ' Si le nom correspond
                    If nameCol.Cells(i, 1).Value = name Then
                        ' Incr�menter le compteur
                        count = count + 1
                        ' Si le nom est trouv� pour la premi�re fois
                        If count = 1 Then
                            ' R�cup�rer le "Folder Path" correspondant
                            getFilePathByName = folderPathCol.Cells(i, 1).Value & name
                        ' Si le nom est trouv� plus d'une fois
                        ElseIf count > 1 Then
                            ' G�n�rer une erreur
                            Err.Raise 1004, , "Nom trouv� plusieurs fois"
                            Exit Function
                        End If
                    End If
                Next i
            End If
        Next tbl
    Next ws

    ' Si le tableau n'est pas trouv�, g�n�rer une erreur
    If Not tableFound Then
        Err.Raise 1004, , "Tableau non trouv�"
    End If

    ' Si le nom n'est pas trouv�, g�n�rer une erreur
    If count = 0 Then
        Err.Raise 1004, , "Nom non trouv�"
    End If
    
End Function
Sub opentest()
    Dim wb As Workbook
    Set wb = openWorkbook("\\stccwp0015\Worksresearsh$\01_Component Librairies\70_NdC et Tol�rances\NX\getExportStringForNX\getExportStringForNX_240312.xlsx", _
        False, False)

End Sub
'Valid� 240328
Private Function openWorkbook(path As String, Optional readOnly As Boolean = False, Optional hidden As Boolean = True) As Workbook
    Dim methodToOpen As String
    'methodToOpen = "Application.Workbooks.Open"
    methodToOpen = "Shell"
    
    If path <> "" Then
        If Dir(path) <> "" Then ' Check if file exists
            If Not IsWorkbookOpen(path) Then ' Check if workbook is already open
                
                Select Case methodToOpen
                    Case "Application.Workbooks.Open"
                        Set openWorkbook = Application.Workbooks.Open(path, , readOnly)

                    Case "Shell"
                        Dim cmd As String
                        cmd = "cmd /c start """" ""excel.exe"" """ & path & """"
                        Call Shell(cmd, vbNormalFocus)
                        
                    Case Else
                End Select
            End If
            ' Wait for the workbook to open
            Set openWorkbook = waitForExcelToOpen(path)
            'MsgBox "excel opened"
                
            If openWorkbook Is Nothing Then
                Err.Raise 1005, , "[openWorkbook] Error: The Object Workbook is nothing with the path: " & path
            Else
                Dim window As Variant
                For Each window In openWorkbook.Windows
                    window.Visible = Not hidden
                Next window
            End If
        Else
            Err.Raise 1007, , "[openWorkbook] Error: The file does not exist at the specified path."
        End If
    Else
        Err.Raise 1008, , "[openWorkbook] Error: No file path provided."
    End If
End Function

'This will way max 10 sec for an excel to open
Private Function waitForExcelToOpen(path As String) As Workbook
    Dim start As Double
    Dim timeout As Double
    Dim workbookName As String
    
    workbookName = Dir(path)
    start = Timer
    timeout = 10 ' Set a timeout (in seconds) to avoid an infinite loop
    
    Do While Not IsWorkbookOpen(workbookName) And Timer - start < timeout
        DoEvents ' Yield to other processes
    Loop
    
    If IsWorkbookOpen(workbookName) Then
        Set waitForExcelToOpen = Application.Workbooks(workbookName)
    Else
        Set waitForExcelToOpen = Nothing
        Err.Raise 1009, , "[waitForExcelToOpen] TimeOut; waited more than " & timeout
    End If
End Function

' Function to check if a workbook is open
Function IsWorkbookOpen(fileName As String) As Boolean

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name = fileName Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
    IsWorkbookOpen = False

End Function

Sub cleanExcelLoaded()
 On Error GoTo ErrorHandler
    Dim win As window
    For Each win In Application.Windows
        If win.Caption = ThisWorkbook.name Then
            ' Ne ferme pas le classeur qui a lanc� la fonction
        Else
            win.Activate
            ActiveWindow.Close saveChanges:=saveModification
        End If
    Next
    Exit Sub
    ErrorHandler:
    MsgBox "Erreur lors du cleanup des excel : " & Err.Description
End Sub




