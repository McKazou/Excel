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
Private Const DEFAULT_LOADING_FOLDER As String = "\\STCCWP0015\pierre.leroy$\07_Code\Excel\Excel Addins Management\"

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
    ' Initialise le projet VBA pour ce classeur
    Set this.VBProject = ThisWorkbook.VBProject
    Me.refreshModule "\\STCCWP0015\pierre.leroy$\07_Code\Excel\File Management\FileManager.cls"
End Sub

Public Sub refreshModule(cheminDuModule As String)
    Dim nomDuModule As String
    Dim dateDeDerniereModification As Date
    Dim managedFile As New FileManager
    managedFile.setThisWorkbook
    
    ' Obtenez le nom du module à partir du chemin du module
    nomDuModule = Mid(cheminDuModule, InStrRev(cheminDuModule, "\") + 1)
    nomDuModule = Left(nomDuModule, InStrRev(nomDuModule, ".") - 1)
    
    ' Obtenez la date de dernière modification du fichier
    dateDeDerniereModification = FileDateTime(cheminDuModule)
    
    ' Si la date de dernière modification est plus récente que la date de dernière sauvegarde du classeur, rechargez le module
    If dateDeDerniereModification > managedFile.lastSaved Then
        ' Fermez le module s'il existe déjà
        Me.CloseModule nomDuModule
        
        ' Importez le module à partir du chemin spécifié
        this.VBProject.VBComponents.Import cheminDuModule
    ' Si la date de dernière sauvegarde du classeur est plus récente que la date de dernière modification du fichier, exportez le module vers le fichier
    ElseIf managedFile.lastSaved > dateDeDerniereModification Then
        ' Fermez le module s'il existe déjà
        Me.CloseModule nomDuModule
        
        ' Exportez le module vers le fichier
        this.VBProject.VBComponents.Item(nomDuModule).Export cheminDuModule
    End If
End Sub




' Méthode pour ajouter un module à partir de son chemin
Public Sub addModule(cheminDuModule As String)
    ' Essayez d'importer un module à partir d'un fichier
    On Error GoTo ErrorHandler
    this.VBProject.VBComponents.Import cheminDuModule
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
        If vbComp.name = nomDuModule Then
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
    moduleActif = Application.VBE.ActiveCodePane.CodeModule.name

    ' Parcourez tous les composants du projet VBA
    For Each vbComp In this.VBProject.VBComponents
        ' Si le composant n'est pas le module actif et est un module standard ou un module de classe, ajoutez-le à la liste des modules à supprimer
        If vbComp.name <> moduleActif And (vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule) Then
            modulesASupprimer.Add vbComp.name
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
        modulesCharges.Add vbComp.name
    Next vbComp

    ' Retournez la collection des modules chargés
    Set modulesCharges = modulesCharges
End Property

' Méthode pour vérifier et recharger les modules si nécessaire
Public Sub ScheduleCheckReloadModules()
    Dim vbComp As VBIDE.VBComponent
    Dim cheminDuModule As String
    Dim dateDeDerniereModification As Date
    Dim managedFile As New FileManager
    Dim modulefile As New FileManager
    managedFile.setThisWorkbook

    ' Parcourez tous les composants du projet VBA
    For Each vbComp In this.VBProject.VBComponents
        ' Ignorez "ThisWorkbook", "XlmaManager" et les feuilles de calcul
        If vbComp.name = "ThisWorkbook" Or vbComp.name = "XlmaManager" Or vbComp.Type = vbext_ct_MSForm Or vbComp.Type = vbext_ct_Document Or vbComp.Type = vbext_ct_StdModule Then GoTo NextComponent

        ' Utilisez FindFile pour rechercher le module dans le dossier DEFAULT_LOADING_FOLDER
        Set modulefile = modulefile.FindFile(vbComp.name, DEFAULT_LOADING_FOLDER, True)
        
        ' Si le module est trouvé, définissez cheminDuModule sur modulefile.FullPath
        If Not modulefile Is Nothing Then
            cheminDuModule = modulefile.FullPath
        Else
            ' Si le module n'est pas trouvé, passez au composant suivant
            GoTo NextComponent
        End If
        
        ' Obtenez la date de dernière modification du fichier
        dateDeDerniereModification = FileDateTime(cheminDuModule)
        
        ' Si la date de dernière modification est plus récente que la date de dernière sauvegarde du classeur, rechargez le module
        If dateDeDerniereModification > managedFile.lastSaved Then
            ' Rafraîchissez le module
            Me.refreshModule cheminDuModule
        ' Si la date de dernière sauvegarde du classeur est plus récente que la date de dernière modification du fichier, exportez le module vers le fichier
        ElseIf managedFile.lastSaved > dateDeDerniereModification Then
            vbComp.Export cheminDuModule
        End If
NextComponent:
    Next vbComp
End Sub










Public Sub test()
    ' Cette méthode est actuellement vide
End Sub

