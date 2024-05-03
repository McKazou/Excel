VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fonction pour vérifier si un tableau existe
Public Function TableExists(TableName As String) As Boolean
    On Error Resume Next
    TableExists = False
    If ThisWorkbook.Worksheets(TableName).ListObjects.Count > 0 Then
        TableExists = True
    End If
End Function

' Fonction pour supprimer un tableau
Public Sub DeleteTable(TableName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(TableName).ListObjects(1).Delete
End Sub

' Fonction pour créer un tableau
Public Sub CreateTable(TableName As String, RangeString As String, IncludeHeaders As Boolean)
    Dim rng As Range
    Dim tbl As ListObject

    ' Définir la plage sur laquelle le tableau sera créé
    Set rng = ThisWorkbook.Range(RangeString)

    ' Créer le tableau
    Set tbl = ThisWorkbook.ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)

    ' Définir le nom du tableau
    tbl.Name = TableName

    ' Spécifier si le tableau doit inclure des en-têtes
    If IncludeHeaders Then
        tbl.HeaderRowRange.EntireRow.Hidden = False
    Else
        tbl.HeaderRowRange.EntireRow.Hidden = True
    End If
End Sub


' Fonction pour obtenir l'objet SuperTable associé à une requête
Public Function GetTableObject(TableName As String) As ListObject
    On Error Resume Next
    Err.Raise 1004, "TableTools", "Not Yet Implemented"
End Function


Public Function getTableNameFromRange(rangeToSearch As Range) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim intersectRange As Range

    ' Parcourir chaque feuille de calcul et chaque tableau
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            ' Vérifier si la plage spécifiée intersecte entièrement avec le tableau
            Set intersectRange = Application.Intersect(tbl.Range, rangeToSearch)
            If Not intersectRange Is Nothing Then
                If intersectRange.Address = tbl.Range.Address Then
                    ' Retourner le nom du tableau
                    getTableFromRange = tbl.Name
                    Exit Function
                End If
            End If
        Next tbl
    Next ws

    ' Si aucun tableau n'a été trouvé, retourner une chaîne vide
    getTableFromRange = ""
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
