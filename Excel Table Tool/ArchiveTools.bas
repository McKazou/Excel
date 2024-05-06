Attribute VB_Name = "ArchiveTools"

Option Explicit

' Propriété pour le préfixe d'archive par défaut
Private Const DEFAULT_ARCHIVE_PREFIXE As String = "HISTORY_"
Private Const DEFAULT_ARCHIVE_SHEET As String = "ARCHIVE_"
Private Const DEFAUT_TIMESTAMPS_COLUMN_NAME As String = "Timestamps"


' Réalise une sauvegarde des tableaux indiqués
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
