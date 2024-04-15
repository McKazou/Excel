VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'test
' Classe FileHandler
Private pPath As String
Private pReadOnly As Boolean
Private pHidden As Boolean
' Propriété fullPath
Private pFullPath As String

' Méthode pour définir le classeur actuel
Public Sub setThisWorkbook()
    ' Définir le chemin du classeur actuel
    pPath = ThisWorkbook.path
    
    ' Définir si le classeur est en lecture seule
    pReadOnly = ThisWorkbook.ReadOnly
    
    ' Vérifier si le classeur est ouvert dans une fenêtre
    If ThisWorkbook.Windows.count > 0 Then
        ' Définir si le classeur est caché
        pHidden = Not ThisWorkbook.Windows(1).Visible
    Else
        ' Si le classeur n'est pas ouvert dans une fenêtre, définir pHidden comme False
        pHidden = False
    End If
    
    ' Définir le chemin complet du classeur
    pFullPath = ThisWorkbook.path & "\" & ThisWorkbook.name
End Sub



' Méthode pour ouvrir un fichier
Public Sub openFile()
    ' Code pour ouvrir un fichier
End Sub

' Méthode pour fermer un fichier
Public Sub closeFile()
    ' Code pour fermer un fichier
End Sub

' Méthode pour vérifier si un fichier existe
Public Function fileExists() As Boolean
    ' Code pour vérifier si un fichier existe
End Function

' Méthode pour obtenir la date de dernière sauvegarde d'un fichier
Public Function lastSaved() As Date
    ' Vérifiez si le fichier existe
    If Dir(pPath) <> "" Then
        ' Obtenez la date de dernière modification du fichier
        lastSaved = FileDateTime(pPath)
    Else
        ' Si le fichier n'existe pas, renvoyez une date vide
        lastSaved = 0
    End If
End Function

' Méthode pour obtenir un fichier à partir d'un chemin
Public Function getFileFromPath(path As String) As String
    ' Code pour obtenir un fichier à partir d'un chemin
End Function

' Méthode pour obtenir un dossier à partir d'un chemin
Public Function getFolderFromPath(path As String) As String
    ' Code pour obtenir un dossier à partir d'un chemin
End Function

' Méthode pour obtenir le dossier où le script est exécuté
Public Function getFolderWhereScriptIsRun() As String
    ' Code pour obtenir le dossier où le script est exécuté
End Function

' Méthode pour obtenir le chemin où le script est exécuté
Public Function getPathWhereScriptIsRun() As String
    ' Code pour obtenir le chemin où le script est exécuté
End Function

Private pFullPath As String

Public Sub FileFinder(FileName As String, selectedFolder As String, isRecursif As Boolean, recursifLimit As Integer)
    Dim strTemp As String
    Dim strFolders() As String
    Dim intFolderCount As Integer
    Dim i As Integer
    
    If recursifLimit = 0 Then Exit Sub
    
    strTemp = Dir(selectedFolder & "\" & FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
    
    While strTemp <> ""
        If strTemp = FileName Then
            pFullPath = selectedFolder & "\" & strTemp
            Exit Sub
        End If
        If (GetAttr(selectedFolder & "\" & strTemp) And vbDirectory) = vbDirectory Then
            If strTemp <> "." And strTemp <> ".." Then
                intFolderCount = intFolderCount + 1
                ReDim Preserve strFolders(1 To intFolderCount)
                strFolders(intFolderCount) = selectedFolder & "\" & strTemp
            End If
        End If
        strTemp = Dir
    Wend
    
    If isRecursif Then
        For i = 1 To intFolderCount
            FileFinder FileName, strFolders(i), True, recursifLimit - 1
            If pFullPath <> "" Then Exit Sub
        Next i
    End If
End Sub


' Méthode pour renommer un fichier
Public Sub renameFileTo()
    ' Code pour renommer un fichier
End Sub

' Méthode pour obtenir le dossier de l'utilisateur
Public Function getUserFolder() As String
    ' Code pour obtenir le dossier de l'utilisateur
End Function

' Méthode pour déplacer un fichier
Public Sub moveFileTo()
    ' Code pour déplacer un fichier
End Sub

