Attribute VB_Name = "ArchiveTools"

Option Explicit

' Propri�t� pour le pr�fixe d'archive par d�faut
Private Const DEFAULT_ARCHIVE_PREFIXE As String = "HISTORY_"
Private Const DEFAULT_ARCHIVE_SHEET As String = "ARCHIVE_"
Private Const DEFAUT_TIMESTAMPS_COLUMN_NAME As String = "Timestamps"


' R�alise une sauvegarde des tableaux indiqu�s
Public Sub testSnapshot()
    Dim SnapTool As New SnapshotTools
    SnapTool.DEFAULT_ARCHIVE_PREFIXE = DEFAULT_ARCHIVE_PREFIXE
    SnapTool.DEFAULT_ARCHIVE_SHEET = DEFAULT_ARCHIVE_SHEET
    SnapTool.DEFAUT_TIMESTAMPS_COLUMN_NAME = DEFAUT_TIMESTAMPS_COLUMN_NAME
    SnapTool.doSnapshot "Backend_AllTableMerge"

    Dim dateOneYearAgo As Date
    dateOneYearAgo = DateAdd("yyyy", -1, Date)
    SnapTool.PurgeAllOlderThan dateOneYearAgo
End Sub

Public Sub testExportCSV()
    Dim SnapTool As New SnapshotTools
    SnapTool.doCSVSnapshot "Backend_AllTableMerge"
End Sub
