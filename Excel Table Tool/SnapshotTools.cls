VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnapshotTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Cette class va permettre de sauvegarder un tableau sous forme d'un autre tableau dans un onglet spécifique


' Propriétés pour le préfixe d'archive par défaut et la feuille d'archive par défaut
Private pDEFAULT_ARCHIVE_PREFIXE As String
Private pDEFAULT_ARCHIVE_SHEET As String

' Constructeur de la classe
Public Sub Class_Initialize()
    On Error GoTo 0
    ' Initialise les objets needed dans cette class
    'Class to load
    Dim classToLoad As Variant
    classToLoad = Array("QueryTools", "SheetsTools", "TableTools")
    'Dim xlma As XlmaManager
    'xlma.Include (classToLoad)
End Sub

' Méthode pour faire un snapshot d'un tableau
Public Sub doSnapshot(NameGiven As String)
    ' Vérifier que pDEFAULT_ARCHIVE_PREFIXE et pDEFAULT_ARCHIVE_SHEET ont une valeur
    If pDEFAULT_ARCHIVE_PREFIXE = "" Or pDEFAULT_ARCHIVE_SHEET = "" Then
        Err.Raise Number:=1004, _
                  Source:="SnapshotTools.DoSnapshot", _
                  Description:="Les valeurs par défaut pour le préfixe d'archive et la feuille d'archive ne sont pas définies."
        Exit Sub
    End If
    
    'Contient le contenu du tableau à copier
    Dim tabRanges As Range
    'Contient l'emplacement où est chargé la query
    Dim QuerRange As Range
    'Contient des méthodes de manipulation de Query
    Dim queryTool As New QueryTools
    'Contient des méthodes de manipulation d'objet tableau au sens Excel
    Dim tabtool As New TableTools
    'Contient the worksheet temporaire créé dans le cas où c'est une query non chargé, il sera delete à la fin de la fonction
    Dim wsh As Worksheet
    
    ' Vérifier si la requête existe
    If queryTool.QueryExists(NameGiven) Then
        ' Vérifier si la requête est chargée
        If Not queryTool.IsQueryLoaded(NameGiven) Then
        'Si elle n'est pas chargé, on va créé un onglet, le charger dans cette onglet afin de copier les valeurs.
            Dim sheetToo As New SheetsTools
            Dim randText As String

            randText = GenerateRandomText(10, True, False, False)
            Set wsh = sheetToo.CreateSheet(randText)
            queryTool.LoadQueryToWorksheet NameGiven, wsh, wsh.Range("A1")
            'Maintenant la Query est chargé dans l'onglet "wsh"
            NameGiven = tabtool.getTableNameFromRange(wsh.Range("A1"))
            Set tabRanges = tabtool.GetContent(NameGiven, wsh)
        Else
        'Si elle est chargé, on va récupéré l'emplacement de la Query, récupéré le tableau ou elle est, copier le contenu et aller le rajouter à un tableau
            Set QuerRange = queryTool.GetQueryLocation(NameGiven)
            Set wsh = QuerRange.Worksheet
            NameGiven = tabtool.getTableNameFromRange(wsh.Range("A1"))
            Set tabRanges = tabtool.GetContent(NameGiven)
        End If
    Else
        'La Query n'existe pas
        'Try with Table
        If tabtool.TableExists(NameGiven) Then
            'Si la table Existe on va récupéré son contenu pour le copier

            Set tabRanges = tabtool.GetContent(NameGiven)
        Else
            Err.Raise 1005, "SnapshotTools", "No Table with the name " & NameGiven & " found."
            Exit Sub
        End If
    End If
    
    'Ici je dois avoir une range dans tabRanges qui contient le contenu du tableau à sauvegarder (copier vers la destination).

    ' On va créer la destination en cherchant le tableau dont le nom est "NameGiven"&"pDEFAULT_ARCHIVE_PREFIXE"
    Dim archiveTableName As String
    archiveTableName = pDEFAULT_ARCHIVE_PREFIXE & NameGiven
    
    ' Si le tableau n'existe pas, on va le créer dans l'onglet dont le nom est "pDEFAULT_ARCHIVE_SHEET"
    Dim sheetTool As New SheetsTools
    Dim archiveSheet As Worksheet
    If Not sheetTool.SheetExists(pDEFAULT_ARCHIVE_SHEET) Then
        Set archiveSheet = sheetTool.CreateSheet(pDEFAULT_ARCHIVE_SHEET)
    Else
        Set archiveSheet = ThisWorkbook.Worksheets(pDEFAULT_ARCHIVE_SHEET)
    End If
    
    ' Si le tableau n'existe pas, on va le créer
    Dim tableTool As New TableTools
    If Not tableTool.TableExists(archiveTableName) Then
        tableTool.CreateTable archiveSheet, archiveTableName, archiveSheet.Range("A1").Address, True
    End If
    
    ' Copier le contenu du tableau à sauvegarder vers la destination
    ' Obtenir l'objet ListObject pour le tableau
    Dim lo As ListObject
    Set lo = archiveSheet.ListObjects(archiveTableName)
    
    'Ajoute à la fin du tableau
    tableTool.addAtEnd NameGiven, tabRanges, archiveSheet
    
    
    '-----------Clean up--------------
    If Not wsh Is Nothing Then
        ' Store the worksheet name
        Dim wshName As String
        wshName = wsh.Name
        
        ' Delete the worksheet
        sheetToo.DeleteSheet wshName
        
        ' Set the worksheet object to Nothing
        Set wsh = Nothing
    End If



End Sub


    
' Méthode pour faire un snapshot de tous les tableaux dans la liste
Public Sub doSnapshotsAll(TableNameList As Variant)
    Dim i As Integer
    
    ' Parcourir chaque nom de tableau dans la liste
    For i = LBound(TableNameList) To UBound(TableNameList)
        ' Appeler la fonction doSnapshot pour chaque nom de tableau
        doSnapshot TableNameList(i)
    Next i
End Sub

    
' Méthode pour purger toutes les données plus anciennes qu'une certaine date
Public Sub PurgeAllOlderThan(pastDate As Date)
    ' Code pour purger toutes les données plus anciennes que la date spécifiée
End Sub

' Getters et setters pour les propriétés
Public Property Get DEFAULT_ARCHIVE_PREFIXE() As String
    DEFAULT_ARCHIVE_PREFIXE = pDEFAULT_ARCHIVE_PREFIXE
End Property

Public Property Let DEFAULT_ARCHIVE_PREFIXE(value As String)
    pDEFAULT_ARCHIVE_PREFIXE = value
End Property

Public Property Get DEFAULT_ARCHIVE_SHEET() As String
    DEFAULT_ARCHIVE_SHEET = pDEFAULT_ARCHIVE_SHEET
End Property

Public Property Let DEFAULT_ARCHIVE_SHEET(value As String)
    pDEFAULT_ARCHIVE_SHEET = value
End Property

