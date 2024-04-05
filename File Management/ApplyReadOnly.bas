Attribute VB_Name = "ApplyReadOnly"
Option Explicit

Sub applyReadOnly()

    Dim dicSent
    'Containt the file name as key
    'FolderPath as item
    Set dicSent = CreateObject("Scripting.Dictionary")
    
        Dim nameRange As Variant
        nameRange = Range("Sent_Status[Name]").Value2
        Dim pathRange As Variant
        pathRange = Range("Sent_Status[Folder Path]").Value2
        Dim SharedStatusRange As Variant
        SharedStatusRange = Range("Sent_Status[SHARED STATUS]").Value2
        'Dim readOnlyRange As Variant
        'readOnlyRange = Range("Shared_status[Attributes.ReadOnly]").Value2

    Dim i As Integer

    Dim fileName, folderPath, sharedStatus As Variant
    Dim fileFullPath As String
    For i = 1 To UBound(nameRange)
        fileName = nameRange(i, 1)
        folderPath = pathRange(i, 1)
        sharedStatus = SharedStatusRange(i, 1)
        fileFullPath = folderPath & fileName & ".prt"
        If sharedStatus = "SENT" Then
            'Set the file to read only
            SetAttr fileFullPath, vbReadOnly
        Else
            SetAttr fileFullPath, vbNormal
        End If

    Next
    
    ActiveWorkbook.RefreshAll
End Sub
