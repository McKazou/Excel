VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub LoadQueryToWorksheet(QueryName As String, LoadDataSheet As Worksheet, StartCell As Range)
    With LoadDataSheet.ListObjects.Add(SourceType:=0, Source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName, _
    Destination:=StartCell).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QueryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .Refresh BackgroundQuery:=False
    End With
End Sub


' Fonction pour vérifier si une requête existe
Public Function QueryExists(QueryName As String) As Boolean
    Dim qry As Object
    On Error Resume Next
    Set qry = ThisWorkbook.Queries(QueryName)
    On Error GoTo 0
    QueryExists = Not qry Is Nothing
End Function

' Fonction pour supprimer une requête
Public Sub DeleteQuery(QueryName As String)
    If QueryExists(QueryName) Then
        ThisWorkbook.Queries(QueryName).Delete
    End If
End Sub

' Fonction pour créer une requête
Public Sub CreateQuery(QueryName As String, MCode As String)
    If Not QueryExists(QueryName) Then
        ThisWorkbook.Queries.Add Name:=QueryName, Formula:=MCode
    End If
End Sub

' Fonction pour obtenir l'objet QueryTable associé à une requête
Public Function GetQueryObject(QueryName As String) As QueryTable
    Dim ws As Worksheet
    Dim lo As ListObject
    
    ' Parcourir toutes les feuilles de calcul du classeur
    For Each ws In ThisWorkbook.Worksheets
        ' Parcourir tous les objets ListObject (tableaux) de la feuille de calcul
        For Each lo In ws.ListObjects
            ' Vérifier si l'objet ListObject est associé à la requête
            If lo.QueryTable Is Not Nothing Then
                If lo.QueryTable.WorkbookConnection.Name = QueryName Then
                    ' Retourner l'objet QueryTable
                    Set GetQueryObject = lo.QueryTable
                    Exit Function
                End If
            End If
        Next lo
    Next ws
End Function


' Fonction pour vérifier si une requête est chargée dans une feuille de calcul
Public Function IsQueryLoaded(QueryName As String) As Boolean
    IsQueryLoaded = Not GetQueryObject(QueryName) Is Nothing
End Function

' Fonction pour supprimer le chargement d'une requête dans une feuille de calcul
Public Sub UnloadQuery(QueryName As String, Optional DeleteTable As Boolean = False)
    Dim qt As QueryTable
    Set qt = GetQueryObject(QueryName)
    
    If Not qt Is Nothing Then
        ' Supprimer le chargement de la requête
        qt.Delete
        ' Supprimer le tableau si DeleteTable est True
        If DeleteTable Then qt.ListObject.Delete
    End If
End Sub





