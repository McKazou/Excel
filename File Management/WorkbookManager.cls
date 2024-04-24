VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private openWorkbooks As Dictionnary ' Dictionnaire pour stocker les workbooks ouverts
Private workbookVisibility As Dictionnary ' Dictionnaire pour stocker l'état d'affichage des workbooks
Private managedFiles As FileTool

' Constructeur de la classe
Public Sub Class_Initialize()
    Set openWorkbooks = New Dictionnary ' Initialisez le Dictionnary ici
    Set workbookVisibility = New Dictionnary ' Initialisez le Dictionnary pour l'état d'affichage
    Set managedFiles = New FileTool
    
    '#Import Dictionnary
    Dim xlma As New XlmaManager
    xlma.include Array("Dictionnary"), ThisWorkbook.VBProject
End Sub


' Define the enumeration for the open methods
Public Enum OpenMethod
    MethodWorkbooksOpen = 1
    MethodShell = 2
End Enum



Private Function openWorkbook(path As String, Optional readOnly As Boolean = False, Optional hidden As Boolean = True, Optional methodToOpen As OpenMethod = MethodWorkbooksOpen) As Workbook
    Dim wb As Workbook
    
    If path <> "" Then
        If Dir(path) <> "" Then ' Check if file exists
            Set wb = GetOpenWorkbook(path)
            If wb Is Nothing Then ' Check if workbook is already open
                Application.EnableEvents = False ' Disable events
                Select Case methodToOpen
                    Case MethodWorkbooksOpen
                        Set wb = Application.Workbooks.Open(path, , readOnly)

                    Case MethodShell
                        ' Utilisez la méthode openFile de FileTool
                        managedFiles.openFile "excel.exe", path
                        
                        Application.Wait (Now + TimeValue("0:00:10"))
                        ' Wait for the workbook to open
                        Set wb = waitForExcelToOpen(path)
                        openWorkbooks.Add wb.name, wb ' Ajoutez le workbook au dictionnaire avec le nom du workbook comme clé
                        
                        Application.EnableEvents = True ' Enable events
                    Case Else
                        Err.Raise 1009, , "[openWorkbook] Error: Invalid open method."
                End Select
            End If
            
            If wb Is Nothing Then
                Err.Raise 1005, , "[openWorkbook] Error: The Object Workbook is nothing with the path: " & path
            Else
            'Handle visibility of all workbooks
                Dim window As Variant
                For Each window In wb.Windows
                    window.Visible = Not hidden
                Next window
                Set openWorkbook = wb
            End If
        Else
            Err.Raise 1007, , "[openWorkbook] Error: The file does not exist at the specified path."
        End If
    Else
        Err.Raise 1008, , "[openWorkbook] Error: No file path provided."
    End If
End Function



'This will way max 10 sec for an excel to open
Private Function waitForExcelToOpen(workbookName As String) As Workbook
    Dim start As Double
    Dim timeout As Double
    Dim elapsedTime As Double
    
    start = Timer
    timeout = 10 ' Set a timeout (in seconds) to avoid an infinite loop
    
    Do
        elapsedTime = Timer - start
        If elapsedTime > timeout Then Exit Do
        On Error Resume Next
        Set waitForExcelToOpen = GetObject(workbookName)
        On Error GoTo 0
        If Not waitForExcelToOpen Is Nothing Then Exit Do
        DoEvents
    Loop
End Function

' Function to check if a workbook is open
Function GetOpenWorkbook(pathOrFileName As String) As Workbook
    Dim FileName As String
    FileName = Dir(pathOrFileName)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name = FileName Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOpenWorkbook = Nothing
End Function

' Ferme un classeur spécifique
' wbName : Le nom du classeur à fermer
' saveChanges : Détermine si les modifications apportées au classeur doivent être enregistrées avant sa fermeture
' alertUser : Détermine si Excel doit afficher des alertes pendant la fermeture du classeur
Private Sub CloseWorkbook(wbName As String, Optional saveChanges As Variant, Optional alertUser As Boolean = True)
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Application.DisplayAlerts = alertUser
        wb.Close saveChanges
        Application.DisplayAlerts = True
        ' Retirez le workbook du dictionnaire
        If openWorkbooks.Exists(wbName) Then
            openWorkbooks.Remove wbName
        End If
    Else
        MsgBox "Workbook " & wbName & " is not currently open.", vbInformation
    End If
End Sub


' Affiche tous les classeurs cachés et stocke leur état d'affichage
Public Sub ShowHiddenWorkbooks()
    Dim key As Variant
    Dim wb As Workbook
    Dim window As Variant
    ' Parcourez tous les classeurs dans le dictionnaire openWorkbooks
    For Each key In openWorkbooks.keys
        Set wb = openWorkbooks.item(key)
        ' Parcourez toutes les fenêtres de chaque classeur
        For Each window In wb.Windows
            ' Stockez l'état d'affichage actuel du workbook
            workbookVisibility.Add wb.name, window.Visible
            ' Si une fenêtre est cachée, la rendre visible
            If Not window.Visible Then
                window.Visible = True
            End If
        Next window
    Next key
End Sub

' Restaure l'état d'affichage de chaque workbook
Public Sub RestablishHiddenState()
    Dim key As Variant
    Dim wb As Workbook
    Dim window As Variant
    ' Parcourez tous les classeurs dans le dictionnaire openWorkbooks
    For Each key In openWorkbooks.keys
        Set wb = openWorkbooks.item(key)
        ' Parcourez toutes les fenêtres de chaque classeur
        For Each window In wb.Windows
            ' Si l'état d'affichage stocké est False, cachez la fenêtre
            If Not workbookVisibility.item(wb.name) Then
                window.Visible = False
            End If
        Next window
    Next key
End Sub


' Ferme tous les classeurs cachés dans le dictionnaire openWorkbooks
' saveChanges : Détermine si les modifications apportées au classeur doivent être enregistrées avant sa fermeture
Public Sub CloseHiddenWorkbooks(Optional saveChanges As Variant)
    Dim key As Variant
    Dim wb As Workbook
    Dim window As Variant
    
    Me.RestablishHiddenState
    ' Parcourez tous les classeurs dans le dictionnaire openWorkbooks
    For Each key In openWorkbooks.keys
        Set wb = openWorkbooks.item(key)
        ' Parcourez toutes les fenêtres de chaque classeur
        For Each window In wb.Windows
            ' Si une fenêtre est cachée, fermez le classeur
            If Not window.Visible Then
                CloseWorkbook wb.name, saveChanges
            End If
        Next window
    Next key
End Sub




