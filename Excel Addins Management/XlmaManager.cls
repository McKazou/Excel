VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlmaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type VBIDE
    VBProject As VBIDE.VBProject
End Type

Private this As VBIDE

' Ajoutez cette constante en haut de votre classe
Private Const DEFAULT_LOADING_FOLDER As String = "\\STCCWP0015\pierre.leroy$\07_Code\Excel\"

' Méthode pour charger une liste de modules
Public Sub include(ParamArray nomsDesModules() As Variant)
    Dim nomDuModule As Variant
    For Each nomDuModule In nomsDesModules
        ' Construit le chemin du fichier en utilisant le dossier par défaut et le nom du module
        Dim cheminDuModule As String
        If Dir(DEFAULT_LOADING_FOLDER & nomDuModule & ".bas") <> "" Then
            cheminDuModule = DEFAULT_LOADING_FOLDER & nomDuModule & ".bas"
        ElseIf Dir(DEFAULT_LOADING_FOLDER & nomDuModule & ".cls") <> "" Then
            cheminDuModule = DEFAULT_LOADING_FOLDER & nomDuModule & ".cls"
        Else
            MsgBox "Impossible de trouver le module " & nomDuModule & " dans le dossier : " & DEFAULT_LOADING_FOLDER, vbCritical, "Module non trouvé"
            Next nomDuModule
        End If
        
        ' Appelle la méthode addModule pour charger le module
        Me.addModule cheminDuModule
End Sub


' Constructeur de la classe
Public Sub Class_Initialize()
    On Error GoTo 0
    ' Initialise le projet VBA pour ce classeur
    Set this.VBProject = ThisWorkbook.VBProject
    Me.refreshModule "\\STCCWP0015\pierre.leroy$\07_Code\Excel\File Management\FileManager.cls"
End Sub

Public Sub refreshModule(cheminDuModule As String, Optional targetProject As VBIDE.VBProject)
    Dim nomDuModule As String
    Dim modulesCharges As Collection
    Dim moduleEstCharge As Boolean
    Dim i As Integer
    Dim lastSavedModule As Date
    Dim lastSavedExcel As Date

    ' Obtenez le nom du module à partir du chemin du module
    nomDuModule = Mid(cheminDuModule, InStrRev(cheminDuModule, "\") + 1)
    nomDuModule = Left(nomDuModule, InStrRev(nomDuModule, ".") - 1)

    ' Obtenez la liste des modules chargés
    Set modulesCharges = Me.modulesLoaded

    ' Vérifiez si le module est déjà chargé
    moduleEstCharge = False
    For i = 1 To modulesCharges.Count
        If modulesCharges.Item(i) = nomDuModule Then
            moduleEstCharge = True
            Exit For
        End If
    Next i

    ' Si le module n'est pas chargé, ajoutez-le et terminez la fonction
    If Not moduleEstCharge Then
            ' Si targetProject n'est pas fourni, utilisez this.VBProject
            If targetProject Is Nothing Then
                Me.addModule cheminDuModule
            Else
                Me.addModule cheminDuModule, targetProject
            End If
        
        Exit Sub
    End If

    ' Obtenez la date de dernière modification du module
    lastSavedModule = FileDateTime(cheminDuModule)
    ' Obtenez la date de dernière sauvegarde du classeur
    lastSavedExcel = ThisWorkbook.BuiltinDocumentProperties("Last Save Time")

    ' Si la date de dernière modification du module est plus récente que la date de dernière sauvegarde du classeur, rechargez le module
    If lastSavedModule > lastSavedExcel Then
        ' Fermez le module s'il existe déjà
        Me.CloseModule nomDuModule

    ' Si la date de dernière sauvegarde du classeur est plus récente que la date de dernière modification du module, exportez le module vers le fichier
    ElseIf lastSavedModule < lastSavedExcel Then
        ' Exportez le module vers le fichier
        this.VBProject.VBComponents.Item(nomDuModule).Export cheminDuModule
    End If
End Sub

' Méthode pour ajouter un module à partir de son chemin
Public Sub addModule(cheminDuModule As String, Optional targetProject As VBIDE.VBProject)
    ' Si targetProject n'est pas fourni, utilisez this.VBProject
    If targetProject Is Nothing Then
        Set targetProject = this.VBProject
    End If
    
    ' Essayez d'importer un module à partir d'un fichier
    On Error GoTo ErrorHandler
    targetProject.VBComponents.Import cheminDuModule
    Exit Sub
ErrorHandler:
    ' Affiche un message si le module ne peut pas être chargé
    MsgBox "Impossible de charger le module à partir du chemin : " & cheminDuModule, vbCritical, "Erreur de chargement du module"
End Sub

' Méthode pour fermer un module spécifique
Public Sub CloseModule(nomDuModule As String)
    Dim vbComp As VBIDE.VBComponent
    
    ' Parcourez tous les composants du projet VBA
    For Each vbComp In this.VBProject.VBComponents
        ' Si le nom du composant correspond au nom du module, supprimez le module
        If vbComp.Name = nomDuModule Then
            this.VBProject.VBComponents.Remove vbComp
            Exit Sub
        End If
    Next vbComp
    
    ' Si le module n'existe pas, affichez un message d'erreur
    MsgBox "Le module " & nomDuModule & " n'existe pas dans ce projet.", vbExclamation
End Sub


' Méthode pour fermer tous les modules sauf le module actif
Public Sub CloseAllExceptActiveModules()
    Dim vbComp As VBIDE.VBComponent
    Dim moduleActif As String
    Dim modulesASupprimer As Collection
    Set modulesASupprimer = New Collection

    ' Obtenez le nom du module actif
    moduleActif = Application.VBE.ActiveCodePane.CodeModule.Name

    ' Parcourez tous les composants du projet VBA
    For Each vbComp In this.VBProject.VBComponents
        ' Si le composant n'est pas le module actif et est un module standard ou un module de classe, ajoutez-le à la liste des modules à supprimer
        If vbComp.Name <> moduleActif And (vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule) Then
            modulesASupprimer.Add vbComp.Name
        End If
    Next vbComp

    ' Parcourez la liste des modules à supprimer et supprimez chaque module
    Dim nomDuModule As Variant
    For Each nomDuModule In modulesASupprimer
        Me.CloseModule (nomDuModule)
    Next nomDuModule
End Sub


' Propriété pour obtenir la liste des modules chargés
Public Property Get modulesLoaded() As Collection
    Dim vbComp As VBIDE.VBComponent
    Dim modulesCharges As New Collection

    ' Parcourez tous les composants du projet VBA
    For Each vbComp In this.VBProject.VBComponents
        ' Ajoutez le nom du module à la collection
        modulesCharges.Add vbComp.Name
    Next vbComp

    ' Retournez la collection des modules chargés
    Set modulesLoaded = modulesCharges
End Property

'Method pour savoir si un module avec un nom spécific existe
Private Function ModuleExists(moduleName As String) As Boolean
    Dim modulesCharges As Collection
    Dim i As Integer

    ' Obtenez la liste des modules chargés
    Set modulesCharges = Me.modulesLoaded

    ' Vérifiez si le module est déjà chargé
    ModuleExists = False
    For i = 1 To modulesCharges.Count
        If modulesCharges.Item(i) = moduleName Then
            ModuleExists = True
            Exit Function
        End If
    Next i
End Function


Public Sub ScheduleCheckReloadModules()
    Dim wb As Workbook
    Dim vbComp As VBIDE.VBComponent
    Dim cheminDuModule As String
    Dim managedFile As New FileManager

    ' Parcourez tous les classeurs ouverts
    For Each wb In Application.Workbooks
        ' Parcourez tous les modules du projet VBA du classeur
        For Each vbComp In wb.VBProject.VBComponents
            ' Ignorez certains composants en fonction de leur nom ou de leur type
            If Not (vbComp.Name = "ThisWorkbook" _
                    Or vbComp.Name = "XlmaManager" _
                    Or vbComp.Type = vbext_ct_MSForm _
                    Or vbComp.Type = vbext_ct_Document _
                    Or vbComp.Type = vbext_ct_StdModule) Then
                
                ' Utilisez FileFinder pour obtenir le chemin du module
                ' Essayez d'abord avec l'extension .bas
                managedFile.FileFinder vbComp.Name & ".bas", DEFAULT_LOADING_FOLDER, True, 10
                cheminDuModule = managedFile.FullPath
                
                ' Si aucun fichier n'a été trouvé, essayez avec l'extension .cls
                If cheminDuModule = "" Then
                    managedFile.FileFinder vbComp.Name & ".cls", DEFAULT_LOADING_FOLDER, True, 10
                    cheminDuModule = managedFile.FullPath
                End If
                
                cheminDuModule = managedFile.FullPath

                ' Actualisez le module
                Me.refreshModule cheminDuModule, wb.VBProject
            End If
        Next vbComp
    Next wb
End Sub

Public Sub renameAllExceptActiveModules()
    ' Rôle :
    ' Cette fonction a pour but de renommer tous les modules dans le projet VBA,
    ' à l'exception de ceux qui sont actuellement actifs (c'est-à-dire en cours d'utilisation).
    '
    ' Problème qu'elle tente de résoudre :
    ' Lors de l'importation d'un module dans un projet VBA, si un module du même nom existe déjà,
    ' Excel renomme automatiquement le module importé en ajoutant un chiffre à la fin du nom.
    ' Cela peut poser problème si vous comptez sur le nom du module pour certaines fonctionnalités de votre code. Par example instancier Dim a as new FileManager
    '
    ' Solution envisagée :
    ' Pour éviter ce problème, cette fonction renomme tous les modules qui ne sont pas actuellement actifs.
    ' Cela libère le nom du module pour être utilisé par le module que vous souhaitez importer.
    '
    ' Comment l'appeler :
    ' Dim xlma as XlmaManager
    ' xlma.renameAllExceptActiveModules
    '
    ' Où l'appeler :
    ' Vous devriez appeler cette fonction à un endroit de votre code où vous êtes sûr qu'aucun des modules que vous ne voulez pas renommer n'est actif.
    ' Cela pourrait être au début de votre code, ou juste avant une opération d'importation de module.
    ' Dans :
    '   Private Sub Workbook_afterSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '       Dim xlma As New XlmaManager
    '       xlma.renameAllExceptActiveModules
    '   End Sub
    '
    ' NOTE : Cette fonction n'a pas encore été implémentée. Le code réel pour renommer les modules doit encore être écrit.

End Sub












Public Sub test()
    ' Cette méthode est actuellement vide
End Sub

