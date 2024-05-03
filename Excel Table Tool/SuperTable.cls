VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuperTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe TableTool
Private pWorkbook As Workbook ' Classeur contenant le tableau
Private pWorksheet As Worksheet ' Feuille de calcul contenant le tableau
Private pRange() As Variant ' Emplacement du tableau
Private pHeaders As Collection ' Liste des en-têtes
Private pmasterKeyColumn As Collection ' Valeurs de la première colonne
Private coreRange() As Variant ' "Cœur" du tableau
Private masterKeyNames As Variant ' Noms possibles pour la MasterKeyColumn

' Initialisation de la classe
Private Sub Class_Initialize()
    Set pHeaders = New Collection
    Set masterKeyColumn = New Collection
    masterKeyNames = Array("iD", "Input_Name", "Name", "Path", "FolderPath", "Parameter Type")
End Sub


' Fonction pour rechercher une valeur dans le tableau
Public Function search(valueToSearch As Variant, inArray As String) As TableTool
    Dim i As Long
    Dim j As Long
    Dim searchTable As New TableTool
    Dim found As Boolean

    ' Rechercher inArray dans les en-têtes
    For i = 1 To pHeaders.Count
        If pHeaders(i) = inArray Then
            ' Si inArray est trouvé, rechercher valueToSearch dans la colonne correspondante
            For j = 1 To masterKeyColumn.Count
                If pRange(j + 1, i).value = valueToSearch Then
                    ' Si valueToSearch est trouvé, ajouter la ligne entière à searchTable
                    'searchTable.Range = pRange(j + 1)
                    'searchTable.Headers (pHeaders)
                    'searchTable.masterKeyColumn.Add masterKeyColumn(j)
                    found = True
                    Exit For
                End If
            Next j
        End If
        If found Then Exit For
    Next i

    ' Si inArray n'est pas trouvé dans les en-têtes, rechercher dans la MasterKeyColumn
    If Not found Then
        For i = 1 To masterKeyColumn.Count
            If masterKeyColumn(i) = inArray Then
                ' Si inArray est trouvé, rechercher valueToSearch dans la ligne correspondante
                For j = 1 To pHeaders.Count
                    If pRange(i + 1, j).value = valueToSearch Then
                        ' Si valueToSearch est trouvé, ajouter la colonne entière à searchTable
                        'searchTable.Range = pRange(j)
                        'searchTable.Headers.Add pHeaders(j)
                        'searchTable.masterKeyColumn = masterKeyColumn
                        found = True
                        Exit For
                    End If
                Next j
            End If
            If found Then Exit For
        Next i
    End If

    ' Si inArray n'est pas trouvé ni dans les en-têtes ni dans la MasterKeyColumn, lever une erreur
    If Not found Then
        Err.Raise 1016, , "La valeur " & inArray & " n'a pas été trouvée dans le tableau."
    End If

    ' Retourner le nouvel objet TableTool
    Set search = searchTable
End Function


' Fonction pour obtenir un tableau à partir de son nom
Public Function fromName(tableName As String, Optional Workbook As Workbook = Nothing, Optional sheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Si aucun classeur n'est spécifié, utiliser le classeur actif
    If Workbook Is Nothing Then
        Set pWorkbook = ThisWorkbook
    Else
        Set pWorkbook = Workbook
    End If

    ' Si aucune feuille de calcul n'est spécifiée, utiliser toutes les feuilles de calcul du classeur
    If sheet Is Nothing Then
        ' Parcourir tous les tableaux de toutes les feuilles de calcul
        For Each ws In pWorkbook.Worksheets
            For Each tbl In ws.ListObjects
                ' Si le nom du tableau correspond à tableName
                If tbl.Name = tableName Then
                    ' Mettre à jour les propriétés de l'objet
                    Set pWorksheet = ws
                    pRange = tbl.Range.value
                    updateTable tbl
                    Exit Function
                End If
            Next tbl
        Next ws
    Else
        ' Parcourir tous les tableaux de la feuille de calcul spécifiée
        Set pWorksheet = sheet
        For Each tbl In pWorksheet.ListObjects
            ' Si le nom du tableau correspond à tableName
            If tbl.Name = tableName Then
                ' Mettre à jour les propriétés de l'objet
                pRange = tbl.Range.value
                updateTable tbl
                Exit Function
            End If
        Next tbl
    End If

    ' Si aucun tableau correspondant n'est trouvé, lever une erreur
    Err.Raise 1012, , "Table " & tableName & " not found in the workbook."
End Function


'' Fonction pour mettre à jour les propriétés de l'objet
Private Sub updateTable(tbl As ListObject)
    Dim i As Long
    Dim j As Long
    Dim hdr As Range
    Dim masterKeyIndex As Variant
    Dim masterKeyCount As Integer

    ' Mettre à jour les en-têtes
    For Each hdr In tbl.HeaderRowRange
        pHeaders.Add hdr.value
    Next hdr

    ' Identifier la MasterKeyColumn
    masterKeyCount = 0
    For i = 1 To pHeaders.Count
        For j = LBound(masterKeyNames) To UBound(masterKeyNames)
            If pHeaders(i) = masterKeyNames(j) Then
                masterKeyIndex = i
                masterKeyCount = masterKeyCount + 1
            End If
        Next j
    Next i

    ' Si aucun nom de MasterKeyColumn n'est trouvé, utiliser la première colonne par défaut
    If masterKeyCount = 0 Then
        MsgBox "Aucun nom dans la liste des noms autorisés n'a été trouvé pour le tableau " & tbl.Name & ". La première colonne sera utilisée par défaut."
        Application.StatusBar = "[TableTool]{updateTable} Aucun nom dans la liste des noms autorisés n'a été trouvé pour le tableau " & tbl.Name & ". La première colonne sera utilisée par défaut."
        masterKeyIndex = 1
    ' Si plusieurs noms de MasterKeyColumn sont trouvés, lever une erreur
    ElseIf masterKeyCount > 1 Then
        Err.Raise 1014, , "Plusieurs noms de MasterKeyColumn ont été trouvés dans les en-têtes."
    End If

    ' Mettre à jour les valeurs de la MasterKeyColumn
    For i = 1 To tbl.ListRows.Count
        masterKeyColumn.Add tbl.DataBodyRange(i, masterKeyIndex).value
    Next i

    ' Mettre à jour le "cœur" du tableau
    coreRange = tbl.Range.Offset(1, masterKeyIndex).Resize(tbl.Range.Rows.Count - 1, tbl.Range.Columns.Count - masterKeyIndex).value
End Sub


' Fonction pour remplacer les valeurs dans certaines colonnes et/ou lignes du tableau
Public Sub replaceValues(newValues As Variant, columnNames As Variant, rowValues As Variant)
    Dim i As Long
    Dim j As Long
    Dim colIndex As Variant
    Dim rowIndex As Variant
    Dim replaceCount As Integer

    ' Initialiser le compteur de remplacements
    replaceCount = 0

    ' Compter le nombre de remplacements qui seraient effectués
    For i = 1 To pHeaders.Count
        If IsEmpty(columnNames) Or IsInArray(pHeaders(i), columnNames) Then
            For j = 1 To masterKeyColumn.Count
                If IsEmpty(rowValues) Or IsInArray(masterKeyColumn(j), rowValues) Then
                    replaceCount = replaceCount + 1
                End If
            Next j
        End If
    Next i

    ' Vérifier si la taille de newValues correspond au nombre de remplacements qui seraient effectués
    If replaceCount <> UBound(newValues) - LBound(newValues) + 1 Then
        Err.Raise 1015, , "La taille de newValues ne correspond pas à la taille de l'intersection spécifiée."
    End If

    ' Réinitialiser le compteur de remplacements
    replaceCount = 0

    ' Parcourir toutes les colonnes du tableau
    For i = 1 To pHeaders.Count
        ' Si aucune colonne n'est spécifiée, ou si le nom de la colonne actuelle est dans columnNames
        If IsEmpty(columnNames) Or IsInArray(pHeaders(i), columnNames) Then
            colIndex = i
            ' Parcourir toutes les lignes du tableau
            For j = 1 To masterKeyColumn.Count
                ' Si aucune ligne n'est spécifiée, ou si la valeur de la ligne actuelle est dans rowValues
                If IsEmpty(rowValues) Or IsInArray(masterKeyColumn(j), rowValues) Then
                    rowIndex = j
                    ' Remplacer la valeur à l'intersection de la colonne et de la ligne par newValues
                    pRange(rowIndex + 1, colIndex) = newValues(replaceCount + 1)
                    replaceCount = replaceCount + 1
                End If
            Next j
        End If
    Next i
End Sub



' Fonction pour extraire certaines lignes et/ou colonnes du tableau
Public Function extract(columnNames As Variant, Optional rowValues As Variant, Optional withHeader As Boolean = False, Optional includeMasterKey As Boolean = False) As TableTool
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colIndex As Variant
    Dim rowIndex As Variant
    Dim extractTable As New TableTool

    ' Parcourir toutes les colonnes du tableau
    For i = 1 To pHeaders.Count
        ' Si aucune colonne n'est spécifiée, ou si le nom de la colonne actuelle est dans columnNames
        If IsEmpty(columnNames) Or IsInArray(pHeaders(i), columnNames) Then
            colIndex = i
            ' Parcourir toutes les lignes du tableau
            For j = 1 To masterKeyColumn.Count
                ' Si aucune ligne n'est spécifiée, ou si la valeur de la ligne actuelle est dans rowValues
                If IsEmpty(rowValues) Or IsInArray(masterKeyColumn(j), rowValues) Then
                    rowIndex = j
                    ' Ajouter la valeur à l'intersection de la colonne et de la ligne au nouvel objet TableTool
                    extractTable.Range(rowIndex + 1, colIndex) = pRange(rowIndex + 1, colIndex)
                End If
            Next j
            ' Ajouter l'en-tête de la colonne au nouvel objet TableTool, si withHeader est True
            If withHeader Then
                extractTable.Headers.Add pHeaders(i)
            End If
        End If
    Next i

    ' Ajouter les valeurs de la MasterKeyColumn au nouvel objet TableTool, si includeMasterKey est True
    If includeMasterKey Then
        For i = 1 To masterKeyColumn.Count
            extractTable.masterKeyColumn.Add masterKeyColumn(i)
        Next i
    End If

    ' Retourner le nouvel objet TableTool
    Set extract = extractTable
End Function

' Fonction pour vérifier si une valeur est dans un tableau
Private Function IsInArray(value As Variant, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function


' Propriété pour obtenir ou définir l'emplacement du tableau
Public Property Get Range() As Variant
    Range = pRange
End Property
Public Property Set Range(value As Range)
    pRange = value.value
End Property

' Propriété pour obtenir ou définir les en-têtes
Public Property Get Headers() As Collection
    Set Headers = pHeaders
End Property
Public Property Set Headers(value As Collection)
    Set pHeaders = value
End Property

' Propriété pour obtenir ou définir les valeurs de la première colonne
Public Property Get masterKeyColumn() As Collection
    Set masterKeyColumn = pmasterKeyColumn
End Property
Public Property Set masterKeyColumn(value As Collection)
    Set pmasterKeyColumn = value
End Property
