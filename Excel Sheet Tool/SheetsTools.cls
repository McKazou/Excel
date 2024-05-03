VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Cette class permet la manipulation des onglets dans un excel

' Fonction pour vérifier si une feuille de calcul existe
Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Fonction pour supprimer une feuille de calcul
Public Sub DeleteSheet(sheetName As String)
    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    Else
        Err.Raise 1004, "SheetsTools", "A sheet with the name " & sheetName & " doesn't exist and cannot be deleted"
    End If
End Sub

' Fonction pour créer une feuille de calcul
Public Function CreateSheet(sheetName As String) As Worksheet
    If Not SheetExists(sheetName) Then
        Set CreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        CreateSheet.Name = sheetName
    Else
        Err.Raise 1005, "SheetsTools", "A sheet with the name already exist " & sheetName
    End If
End Function


' Fonction pour obtenir le nom de la feuille de calcul à partir d'une plage de cellules
Public Function GetSheetNameFromRange(rng As Range) As String
    GetSheetNameFromRange = rng.Worksheet.Name
End Function

' Fonction pour obtenir la liste des noms de toutes les feuilles de calcul du classeur
Public Function GetSheetNamesList() As Variant
    Dim ws As Worksheet
    Dim SheetNames() As String
    Dim i As Integer

    i = 0

    ' Parcourir chaque feuille de calcul
    For Each ws In ThisWorkbook.Worksheets
        ' Ajouter le nom de la feuille de calcul au tableau
        ReDim Preserve SheetNames(i)
        SheetNames(i) = ws.Name
        i = i + 1
    Next ws

    ' Retourner le tableau de noms de feuilles de calcul
    GetSheetNamesList = SheetNames
End Function

Public Sub TestClassSheets()
    Dim sheetName As String
    Dim rng As Range
    Dim SheetNames() As String
    
    ' Définir les valeurs de test
    sheetName = "MaFeuille"
    
    ' Tester la fonction CreateSheet
    CreateSheet sheetName
    Debug.Assert SheetExists(sheetName)
    
    ' Tester la fonction GetSheetNameFromRange
    Set rng = ThisWorkbook.Worksheets(sheetName).Range("A1")
    Debug.Assert GetSheetNameFromRange(rng) = sheetName
    
    ' Tester la fonction GetSheetNamesList
    SheetNames = GetSheetNamesList()
    Debug.Assert UBound(SheetNames) >= 0
    Debug.Assert IsInArray(sheetName, SheetNames)
    
    ' Tester la fonction DeleteSheet
    DeleteSheet sheetName
    Debug.Assert Not SheetExists(sheetName)
End Sub

' Fonction auxiliaire pour vérifier si une valeur est dans un tableau
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo ErrorHandler
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
ErrorHandler:
    On Error GoTo 0
End Function

