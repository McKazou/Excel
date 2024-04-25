VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionnaryND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This whole class is to have a multi dimensionnal Dictionnary
'It could have been done using recursive dictionnary
'It also could have been done by "merging" keys together to get one key
'Bilan => could have avoided all this ...

Private nKeys As Collection ' Collection pour stocker les collections de clés
Private pValues As Collection ' Collection pour stocker les valeurs
Private pDimension  As Integer ' Le nombre de dimensions du dictionnaire
Private pCount As Integer

' Initialise les collections lors de la création de l'objet DictionaryND
Private Sub Class_Initialize()
    Set nKeys = New Collection
    Set pValues = New Collection
End Sub

'Obtient la valeur à l'indice i
'indice :           Indice où se situe la valeur à retourner
'Retourne :         La valeur à l'indice donné
Public Property Get values(i As Integer)
    values = pValues(i)
End Property

'Obtient la valeur à l'indice i
'indice :           Indice où se situe les clés à récupéré
'Retourne :         Les clés à l'indice donné
Public Property Get Keys(i As Integer) As Collection
    Dim keysCollection As Collection
    For Each dimension In nKeys
        keysCollection.Add dimension(i)
    Next
    values = keysCollection
End Property

'Obtient le nombre de valeur stocké dans le dictionnaire
'Retourne :         Le nombre de valeurs dans le Dictionnaire à n Dimension
Public Property Get Count() As Integer
    Count = pValues.Count
End Property

' Définit la valeur de la propriété Dimension
Public Property Let dimension(n As Integer)
    Dim i As Integer
    ' Si la dimension a déjà été définie, lève une erreur
    If pDimension <> 0 Then
        Err.Raise 1010, , "Impossible de redimensionner le dictionnaire."
    Else
        If pDimension < 0 Then
            Err.Raise 1011, , "Impossible de créer un dictionaire avec une dimension négative."
        Else
            pDimension = n
            ' Crée une nouvelle collection de clés pour chaque dimension
            For i = 1 To pDimension
                nKeys.Add New Collection
            Next i
        End If
    End If
End Property


' Obtient la valeur de la propriété Dimension
Public Property Get dimension() As Integer
    dimension = pDimension
End Property

' Trouve l'indice du premier ensemble de clés qui correspond aux clés fournies
' keys :            Un tableau de clés à rechercher
' Retourne :        L'indice du premier ensemble de clés qui correspond aux clés fournies. Si aucune clé n'est trouvée, retourne -1.
Private Function FindKeys(Keys() As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim hasAllKey As Boolean
    'Parcours toutes les clés
    For i = 1 To nKeys(1).Count
        hasAllKey = True
        'Parcours toutes les dimensions
        For j = 1 To Me.dimension
            'Test si la clé fourni est vide ou Nothing
            If Keys(j) <> "" Then
                'On calcul si la clé match parmis les nKeys on obtiens "1" sinon "0".
                'Le produit de ces valeurs permets de faire un "ET" logique et de savoir si toutes les clés match
                hasAllKey = (IIf(nKeys(j)(i) = Keys(j), 1, 0)) * hasAllKey
                If hasAllKey = False Then
                    Exit For
                End If
            End If
        Next j
        'Si on a trouvé toutes les clés alors on renvoie l'indice
        If hasAllKey Then
            FindKeys = i
            Exit Function
        End If
    Next i
    'Si aucune clé n'a été trouvée, on renvoie Nothing
    'Err.Raise 1013, , "No Key-Pairs have been found"
    FindKeys = -1
End Function

' Trouve tous les indices des ensembles de clés qui correspondent aux clés fournies
' keys :            Un tableau de clés à rechercher
' Retourne :        Une collection contenant tous les indices des ensembles de clés qui correspondent aux clés fournies.
'                   Si aucune clé n'est trouvée, retourne une collection vide.
Private Function FindAllKeys(Keys() As Variant) As Collection
    Dim i As Integer
    Dim j As Integer
    Dim hasKey As Boolean
    Dim indices As New Collection
    'Parcours toutes les clés
    For i = 1 To nKeys(1).Count
        hasKey = True
        'Parcours toutes les dimensions
        For j = 1 To Me.dimension
            'Test si la clé fourni est vide ou Nothing
            If Keys(j) <> "" Then
                'On calcul si la clé match parmis les nKeys on obtiens "1" sinon "0".
                'Le produit de ces valeurs permets de faire un "ET" logique et de savoir si toutes les clés match
                hasKey = (IIf(nKeys(j)(i) = Keys(j), 1, 0)) * hasKey
                If hasKey = False Then
                    Exit For
                End If
            End If
        Next j
        'Si on a trouvé toutes les clés alors on ajoute l'indice à la collection
        If hasKey Then
            indices.Add i
        End If
    Next i
    'Retourne la collection d'indices
    Set FindAllKeys = indices
End Function



' Ajoute une nouvelle paire clé-valeur au dictionnaire
' value :           La valeur à ajouter
' keys :            Un tableau de clés à ajouter
' Comportement :    Dans le cas ou la clé est déjà présente dans le dictionnaire, pour le moment une erreur est levé, il est possible de "remplacer" la valeur à la place
Public Sub Add(value As Variant, Keys() As Variant)
    Dim i, j, k As Integer
    Dim hasAllKey As Boolean
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        ' Initialise le dictionnaire avec les dimensions de cette première sé
        Me.dimension = UBound(Keys) - LBound(Keys) + 1
    End If
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(Keys) - LBound(Keys) + 1 <> dimension Then
        Err.Raise 1013, , "Number of keys does not match the number of dimensions."
    Else
        ' Vérifie si les clés existent déjà
        i = FindKeys(Keys)
        
        'Si on a trouvé toutes les clés alors on lance une erreur
        If i <> -1 Then
            ' Si les clés existent déjà, on remplace la valeur associée
            Err.Raise 1014, , "Keys " & Join(Keys, ", ") & " already exist in the dictionary."
            Exit Sub
        Else
            ' Ajoute les clés et la valeur si les clés n'existent pas déjà
            For k = 1 To Me.dimension
                nKeys(k).Add Keys(k)
            Next k
            pValues.Add value
        End If
    End If
End Sub


Public Function Exists(Keys() As Variant) As Boolean
    Dim i As Variant
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Utilise la fonction FindKeys pour vérifier l'existence de la clé
        i = FindKeys(Keys)
        If i <> -1 Then
            ' Si les clés existent déjà, retourne True
            Exists = True
            Exit Function
        End If
    End If
    ' Si la clé n'est pas trouvée, retourne False
    Exists = False
End Function


' Vérifie si une clé est dans une collection
' col :             La collection à vérifier
' key :             La clé à vérifier
' Retourne :        True si la clé est dans la collection, False sinon
Private Function IsInCollection(col As Collection, key As Variant) As Boolean
    Dim obj As Variant
    On Error Resume Next
    Set obj = col(key)
    If Err.Number = 0 Then
        IsInCollection = True
    Else
        IsInCollection = False
    End If
    On Error GoTo 0
End Function

' Compare deux dictionnaires pour vérifier s'ils sont identiques
' dict1, dict2 : Les deux dictionnaires à comparer
' Retourne : True si les dictionnaires sont identiques, False sinon
Public Function equal(dict2 As DictionnaryND) As Boolean
    Dim i As Integer
    Dim Keys() As Variant

    ' Vérifie si les deux dictionnaires ont la même dimension
    If Me.dimension <> dict2.dimension Then
        equal = False
        Exit Function
    End If

    ' Vérifie si les deux dictionnaires ont le même nombre de clés
    If Me.Count <> dict2.Count Then
        equal = False
        Exit Function
    End If

    ' Parcourt toutes les clés du premier dictionnaire
    For i = 1 To Me.Count
        If Not Me.values(i) = dict2.values(i) Then
            equal = False
            Exit Function
        End If
    Next i

    ' Si toutes les clés et valeurs sont les mêmes, les dictionnaires sont identiques
    equal = True
End Function


' Supprime une paire cl?-valeur du dictionnaire
' keys :            Un tableau de cl?s de la paire ? supprimer
Public Function Remove(Keys() As Variant)
    Dim i As Variant
    Dim Dimen As Collection
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(Keys) - LBound(Keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Utilise la fonction FindKeys pour trouver l'indice des clés
            i = FindKeys(Keys)
            If i <> -1 Then
                ' Si les clés existent déjà, on les supprime
                pValues.Remove (i)
                For Each Dimen In nKeys
                    Dimen.Remove (i)
                Next
                Exit Function
            End If
        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        Err.Raise 1013, , "Keys " & Join(Keys, ", ") & " not found in the dictionary."
    End If
    On Error GoTo 0
End Function




' Obtient la valeur associée à une paire de clés
' keys :            Un tableau de clés dont la valeur doit être obtenue
' Retourne :        La valeur associée aux clés
Public Function item(Keys() As Variant) As Variant
    Dim i As Variant
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(Keys) - LBound(Keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Utilise la fonction FindKeys pour trouver l'indice des clés
            i = FindKeys(Keys)
            If i <> -1 Then
                ' Si les clés existent déjà, on retourne la valeur
                item = pValues(i)
                Exit Function
            End If
        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        Err.Raise 1013, , "Keys " & Join(Keys, ", ") & " not found in the dictionary."
    End If
    On Error GoTo 0
End Function


' Obtient toutes les valeurs associées à une clé donnée
' keys :            Un tableau de clés dont les valeurs doivent être obtenues
' keepKey :         Un paramètre optionnel pour déterminer si la clé doit être conservée dans le nouveau dictionnaire
' Retourne :        Un objet DictionnaryND contenant les valeurs associées aux clés
Public Function items(Keys() As Variant, Optional keepKey As Boolean = False) As DictionnaryND
    Dim i As Variant
    Dim j, k As Integer
    Dim newKeys() As Variant
    Dim newDict As New DictionnaryND
    Dim indices As Collection
    
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(Keys) - LBound(Keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else

            ' Utilise la fonction FindAllKeys pour trouver tous les indices des clés
            Set indices = FindAllKeys(Keys)
            ' Parcourt tous les indices trouvés
            For Each i In indices
                k = 1
                ' Crée un nouveau tableau de clés pour le nouveau dictionnaire
                If keepKey Then
                     ReDim newKeys(1 To Me.dimension)
                Else
                    ReDim newKeys(1 To Me.dimension - 1)
                End If
                
                For j = 1 To Me.dimension
                    If Keys(j) = "" Then
                        newKeys(k) = nKeys(j)(i)
                        k = k + 1
                    Else
                        If keepKey Then
                            newKeys(k) = nKeys(j)(i)
                            k = k + 1
                        Else
                            'skip this one
                        End If
                    End If
                Next j
                ' Ajoute la valeur au nouveau dictionnaire
                newDict.Add pValues(i), newKeys
            Next i
        End If
    End If
    ' Retourne le nouveau dictionnaire
    Set items = newDict
End Function






' Crée un tableau 2D à partir du dictionnaire
' Retourne :        Un tableau 2D où chaque ligne est une paire clé-valeur du dictionnaire
Public Function toArray() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim arr() As Variant
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        Err.Raise 1011, , "Dimension of the dictionary has not been specified."
    Else
        ' Redimensionne le tableau pour contenir toutes les paires clé-valeur
        ReDim arr(1 To nKeys(1).Count, 1 To dimension + 1)
        ' Remplit le tableau avec les clés et les valeurs
        For i = 1 To nKeys(1).Count
            For j = 1 To dimension
                arr(i, j) = nKeys(j)(i)
            Next j
            arr(i, dimension + 1) = pValues(i)
        Next i
    End If
    ' Retourne le tableau
    toArray = arr
End Function

' Initialise le dictionnaire à partir d'un tableau 2D
' arr :         Un tableau 2D où chaque ligne est une paire clé-valeur
Public Function FromArray(arr() As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim newKeys As Collection
    ' Vérifie si la dimension du tableau a été spécifiée
    If UBound(arr, 2) = 0 Then
        Err.Raise 1011, , "Dimension of the array has not been specified."
    Else
        ' Redimensionne le dictionnaire pour contenir toutes les paires clé-valeur
        Me.dimension = UBound(arr, 2) - 1

        Set pValues = New Collection
        ' Remplit le dictionnaire avec les clés et les valeurs
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2) - 1
                nKeys(j).Add arr(i, j)
            Next j
            pValues.Add arr(i, UBound(arr, 2))
        Next i
    End If
End Function


Sub Test()
    Dim dict As New DictionnaryND
    Dim dictFromArray As New DictionnaryND
    Dim Keys(1 To 2) As Variant
    Dim value As Variant
    Dim i, j As Integer
    
    ' Initialise le dictionnaire avec 2 dimensions
    dict.dimension = 2
    
    ' Ajoute des paires clé-valeur au dictionnaire
    For i = 1 To 3
        For j = 1 To 3
            Keys(1) = "Key1_" & i
            Keys(2) = "Key2_" & j
            value = "Value_" & i & j
            dict.Add value, Keys
        Next j
    Next i
    
    
    
    '-------------------------------------------
    'Vérification getting a specific value : Valid function 'item'
    '-------------------------------------------
    Dim result As String
    result = dict.item(Keys)
    If Not result = value Then
        Debug.Print "ERROR - Got the item :" & result & " for the keys {" & Keys(1) & "," & Keys(2) & "} should have gotten " & value
    Else
        Debug.Print "VALID - Got the item :" & result & " for the keys {Key1_10, Key2_10} - VALID"
    End If
    
    
    
    '-------------------------------------------
    'Vérification a value has been removed : Valid function 'Exists' & 'Remove'
    '-------------------------------------------
    
    ' Supprime des paires clé-valeur du dictionnaire
    Keys(1) = "Key1_" & 1
    Keys(2) = "Key2_" & 1
    dict.Remove Keys
    
    If dict.Exists(Keys) Then
        Debug.Print "ERROR - The item has not been removed correctly from the dictionnary - ERROR"
    Else
        Debug.Print "VALID - The item has been removed correctly from the dictionnary - VALID"
    End If
    
    
    
    
    '-------------------------------------------
    'Vérification que dict et dictFromArray sont identiques : Valid function 'equal', 'toArray' & 'FromArray'
    '-------------------------------------------
    'Initialisation des tableaux de sorti
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim arr3() As Variant
    
    arr = dict.toArray 'Convert to array
    
    dictFromArray.FromArray arr 'Create from Array

    arr2 = dictFromArray.toArray
    
    If Not dict.equal(dictFromArray) Then
        Debug.Print "ERROR - Array are not identical - ERROR"
    Else
        Debug.Print "VALID - Array are identical - VALID"
    End If
    
    

    
    
    '-------------------------------------------
    'Verification we did get a slice of the dictionnary : Valid function 'items'
    '-------------------------------------------
    Dim slice As DictionnaryND
    Keys(1) = "Key1_1"
    Keys(2) = ""
    Set slice = dictFromArray.items(Keys) ' doesn't return the dict
    arr3 = slice.toArray
    
    MsgBox "End of test"
    
End Sub

