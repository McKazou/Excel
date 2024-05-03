VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Cette Class permet de manipuler facilement les tableaux


' Fonction pour obtenir la plage de cellules d'un tableau
Public Function GetContent(tableName As String, Optional ws As Worksheet) As Range
    Dim lo As ListObject
    
    ' Si ws n'est pas défini, utiliser toutes les feuilles de calcul du classeur
    If ws Is Nothing Then
        Dim w As Worksheet
        For Each w In ThisWorkbook.Worksheets
            For Each lo In w.ListObjects
                If lo.Name = tableName Then
                    Set GetContent = lo.Range
                    Exit Function
                End If
            Next lo
        Next w
    Else
        ' Sinon, utiliser uniquement la feuille de calcul spécifiée
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set GetContent = lo.Range
                Exit Function
            End If
        Next lo
    End If
    
    ' Si aucun tableau n'a été trouvé, retourner Nothing
    Set GetContent = Nothing
End Function


' Fonction pour vérifier si un tableau existe
Public Function TableExists(tableName As String, Optional ws As Worksheet) As Boolean
    Dim lo As ListObject
    
    ' Si ws n'est pas défini, utiliser toutes les feuilles de calcul du classeur
    If ws Is Nothing Then
        Dim w As Worksheet
        For Each w In ThisWorkbook.Worksheets
            For Each lo In w.ListObjects
                If lo.Name = tableName Then
                    TableExists = True
                    Exit Function
                End If
            Next lo
        Next w
    Else
        ' Sinon, utiliser uniquement la feuille de calcul spécifiée
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                TableExists = True
                Exit Function
            End If
        Next lo
    End If
    
    ' Si le tableau n'est pas trouvé, retourner False
    TableExists = False
End Function



' Fonction pour supprimer un tableau
Public Sub DeleteTable(tableName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(tableName).ListObjects(1).Delete
End Sub

' Fonction pour créer un tableau
Public Sub CreateTable(wsheet As Worksheet, tableName As String, RangeString As String, IncludeHeaders As Boolean)
    Dim rng As Range
    Dim tbl As ListObject

    ' Définir la plage sur laquelle le tableau sera créé
    Set rng = wsheet.Range(RangeString)

    ' Créer le tableau
    Set tbl = wsheet.ListObjects.Add(xlSrcRange, rng, , IncludeHeaders)

    ' Définir le nom du tableau
    tbl.Name = tableName
End Sub


' Function to get the name of tha table that intersect with the position of the range
Public Function getTableNameFromRange(rangeToSearch As Range) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim intersectRange As Range
    Set ws = rangeToSearch.Worksheet
        For Each tbl In ws.ListObjects
            ' Vérifier si la plage spécifiée est sur la même feuille de calcul que le tableau
            If rangeToSearch.Worksheet.Name = ws.Name Then
                ' Vérifier si la plage spécifiée intersecte entièrement avec le tableau
                Set intersectRange = Application.Intersect(tbl.Range, rangeToSearch)
                If Not intersectRange Is Nothing Then
                        ' Retourner le nom du tableau
                        getTableNameFromRange = tbl.Name
                        Exit Function
                End If
            End If
        Next tbl

    ' Si aucun tableau n'a été trouvé, retourner une chaîne vide
    getTableNameFromRange = ""
End Function


'Get the list of table présent in the workbook
Function getTablesNamesList() As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim TableNames() As String
    Dim i As Integer

    i = 0

    ' Parcourir chaque feuille de calcul et chaque tableau
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            ' Ajouter le nom du tableau au tableau
            ReDim Preserve TableNames(i)
            TableNames(i) = tbl.Name
            i = i + 1
        Next tbl
    Next ws

    ' Retourner le tableau de noms de tableaux
    getTablesNamesList = TableNames
End Function

' Fonction pour ajouter des données à la fin d'un tableau
Public Sub addAtEnd(tableName As String, rangeToPaste As Range, Optional ws As Worksheet)
    Dim tbl As ListObject
    Dim rng As Range

    ' Si aucune feuille de calcul n'est spécifiée, utilisez la feuille de calcul active
    If ws Is Nothing Then Set ws = ActiveSheet

    ' Vérifiez si le tableau existe
    If Not TableExists(tableName, ws) Then
        MsgBox "Table " & tableName & " does not exist."
        Exit Sub
    End If

    ' Obtenez le tableau et la plage de cellules
    Set tbl = ws.ListObjects(tableName)
    Set rng = tbl.Range

    ' Ajoutez les données à la fin du tableau
    rng.Resize(rng.Rows.Count + rangeToPaste.Rows.Count, rng.Columns.Count).value = rangeToPaste.value
End Sub


Public Sub TestTableTools()
    Dim tableName As String
    Dim RangeString As String
    Dim IncludeHeaders As Boolean
    Dim rng As Range
    Dim TableNames() As String
    
    ' Définir les valeurs de test
    tableName = "MonTableau"
    RangeString = "A1:C10"
    IncludeHeaders = True
    
    ' Tester la fonction CreateTable
    CreateTable tableName, RangeString, IncludeHeaders
    Debug.Assert TableExists(tableName)
    
    ' Tester la fonction GetContent
    Set rng = GetContent(tableName)
    Debug.Assert Not rng Is Nothing
    Debug.Print "Le tableau est dans la plage : " & rng.Address
    
    ' Tester la fonction getTableNameFromRange
    Debug.Assert getTableNameFromRange(rng) = tableName
    
    ' Tester la fonction getTablesNamesList
    TableNames = getTablesNamesList()
    Debug.Assert UBound(TableNames) >= 0
    Debug.Assert IsInArray(tableName, TableNames)
    
    ' Tester la fonction DeleteTable
    DeleteTable tableName
    Debug.Assert Not TableExists(tableName)
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

