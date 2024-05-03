Attribute VB_Name = "ArchiveTools"

Option Explicit

' Propriété pour le préfixe d'archive par défaut
Private Const DEFAULT_ARCHIVE_PREFIXE As String = "HISTORY_"
Private Const DEFAULT_ARCHIVE_SHEET As String = "ARCHIVE_"


' Réalise une sauvegarde des tableaux indiqués
Public Sub testSnapshot()
    Dim ArchiveTableList As Variant
    ArchiveTableList = Array("Backend_Parts_List_Data", "Frontend_AllTableMerge")
    
    Dim SnapTool As New SnapshotTools
    SnapTool.DEFAULT_ARCHIVE_PREFIXE = DEFAULT_ARCHIVE_PREFIXE
    SnapTool.DEFAULT_ARCHIVE_SHEET = DEFAULT_ARCHIVE_SHEET
    SnapTool.doSnapshot "Backend_Parts_List_Data"
    SnapTool.doSnapshot "Frontend_Bill_Of_Materials"
    
    SnapTool.doSnapshotsAll ArchiveTableList
    
    Dim dateOneYearAgo As Date
    dateOneYearAgo = DateAdd("yyyy", -1, Date)
    SnapTool.PurgeAllOlderThan dateOneYearAgo
End Sub
