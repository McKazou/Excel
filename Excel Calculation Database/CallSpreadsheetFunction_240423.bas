Attribute VB_Name = "CallSpreadsheetFunction_240423"

Private SpreadsheetFunctionList As Object
Public DEFAULT_TABLE_CALCULATION_DATABASE_NAME As String

Private managedWorkbooks As New WorkbookManager


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

Function CallFunction(functionName As String, Optional inputRange As Range, Optional overrideRange As Range) As String
    Call Init
    ' Création d'une instance de l'objet tableTool
    Dim functionTable As New tableTool
    
    ' Utilisation de la méthode getTableFromName pour obtenir le tableau
    functionTable.fromName (DEFAULT_TABLE_CALCULATION_DATABASE_NAME)
    
    ' Recherche de la valeur dans le tableau
    Dim filePath As New tableTool
    Set filePath = functionTable.search(functionName, "Name")
    
    'Open the file with the path
    Dim wb As Workbook
    Set wb = managedWorkbooks.openWorkbook(filePath, readOnly = False, hidden = True)
    
    Dim InputTable As tableTool
    InputTable.fromName "Input", wb
    InputTable.replaceValues inputRange, "Value"
    
    Dim overrideTable As tableTool
    overrideTable.fromName "Input", wb
    overrideTable.replaceValues overrideRange, "Override_value"
    
    wb.Calculate
    
    Dim outputTable As tableTool
    InputTable.fromName "Output", wb
    
    CallFunction = outputTable.extract(Array("Parameter_Name_In_Calculation", "Value"), , True)
    
End Function




