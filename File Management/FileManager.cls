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
Private pFullPath As String

' Propriété Path
Public Property Let Path(ByVal Value As String)
    pPath = Value
End Property

Public Property Get Path() As String
    Path = pPath
End Property

' Propriété ReadOnly
Public Property Let ReadOnly(ByVal Value As Boolean)
    pReadOnly = Value
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = pReadOnly
End Property

' Propriété Hidden
Public Property Let Hidden(ByVal Value As Boolean)
    pHidden = Value
End Property

Public Property Get Hidden() As Boolean
    Hidden = pHidden
End Property

' Propriété FullPath
Public Property Let FullPath(ByVal Value As String)
    pFullPath = Value
End Property

Public Property Get FullPath() As String
    FullPath = pFullPath
End Property

' Méthode pour définir le classeur actuel
Public Sub setThisWorkbook()
    ' Définir le chemin du classeur actuel
    pPath = ThisWorkbook.Path
    
    ' Définir si le classeur est en lecture seule
    pReadOnly = ThisWorkbook.ReadOnly
    
    ' Vérifier si le classeur est ouvert dans une fenêtre
    If ThisWorkbook.Windows.Count > 0 Then
        ' Définir si le classeur est caché
        pHidden = Not ThisWorkbook.Windows(1).Visible
    Else
        ' Si le classeur n'est pas ouvert dans une fenêtre, définir pHidden comme False
        pHidden = False
    End If
    
    ' Définir le chemin complet du classeur
    pFullPath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
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
    If pPath <> "" Then
        ' Obtenez la date de dernière modification du fichier
        lastSaved = FileDateTime(pPath)
    Else
        ' Si le fichier n'existe pas, renvoyez une date vide
        lastSaved = 0
    End If
End Function

' Méthode pour obtenir un fichier à partir d'un chemin
Public Function getFileFromPath(Path As String) As String
    ' Code pour obtenir un fichier à partir d'un chemin
End Function

' Méthode pour obtenir un dossier à partir d'un chemin
Public Function getFolderFromPath(Path As String) As String
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

Public Sub FileFinder(FileName As String, selectedFolder As String, isRecursif As Boolean, recursifLimit As Integer, Optional ignoreHidden As Boolean = True)
    Dim FSO As Object
    Dim Folder As Object
    Dim Subfolder As Object
    Dim File As Object

    If recursifLimit = 0 Then Exit Sub

    ' Vérifiez si selectedFolder se termine par un backslash
    If Right(selectedFolder, 1) <> "\" Then
        selectedFolder = selectedFolder & "\"
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(selectedFolder)

    ' Parcourir tous les fichiers dans le dossier
    For Each File In Folder.Files
        If File.Name = FileName Then
            pFullPath = File.Path
            pPath = selectedFolder
            Exit Sub
        End If
    Next File

    ' Parcourir les sous-dossiers si isRecursif est True
    If isRecursif Then
        For Each Subfolder In Folder.Subfolders
            ' Ignore les dossiers cachés si ignoreHidden est True
            If ignoreHidden And Subfolder.Attributes And 2 Then
                ' Ce dossier est caché, donc on l'ignore
            Else
                FileFinder FileName, Subfolder.Path, isRecursif, recursifLimit - 1, ignoreHidden
            End If
        Next Subfolder
    End If

    Set File = Nothing
    Set Folder = Nothing
    Set FSO = Nothing
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

