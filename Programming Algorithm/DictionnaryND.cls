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

Private nKeys As Collection ' Collection pour stocker les collections de clés
Private values As Collection ' Collection pour stocker les valeurs
Private pDimension  As Integer ' Le nombre de dimensions du dictionnaire

' Initialise les collections lors de la création de l'objet DictionaryND
Private Sub Class_Initialize()
    Set nKeys = New Collection
    Set values = New Collection
End Sub

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

Private Function FindKeys(keys() As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim hasAllKey As Boolean
    'Parcours toutes les clés
    For i = 1 To nKeys(1).count
        hasAllKey = True
        'Parcours toutes les dimensions
        For j = 1 To Me.dimension
            'Test si la clé fourni est vide ou Nothing
            If keys(j) <> Nothing Or keys(j) <> "" Then
                'On calcul si la clé match parmis les nKeys on obtiens "1" sinon "0".
                'Le produit de ces valeurs permets de faire un "ET" logique et de savoir si toutes les clés match
                hasAllKey = (IIf(nKeys(j)(i) = keys(j), 1, 0)) * hasAllKey
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
    FindKeys = Nothing
End Function


' Ajoute une nouvelle paire clé-valeur au dictionnaire
' value : La valeur à ajouter
' keys : Un tableau de clés à ajouter
' Comportement : Dans le cas ou la clé est déjà présente dans le dictionnaire, pour le moment une erreur est levé, il est possible de "remplacer" la valeur à la place
Public Sub Add(value As Variant, keys() As Variant)
    Dim i, j, k As Integer
    Dim hasAllKey As Boolean
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        ' Initialise le dictionnaire avec les dimensions de cette première sé
        Me.dimension = UBound(keys) - LBound(keys) + 1
    End If
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(keys) - LBound(keys) + 1 <> dimension Then
        Err.Raise 1013, , "Number of keys does not match the number of dimensions."
    Else
        'Parcours toutes les clés
            For i = 1 To nKeys(1).count
                
                ' Vérifie si les clés existent déjà
                i = FindKeys(keys)
                
                'Si on a trouvé toutes les clés alors on lance une erreur
                If i <> Nothing Then
                    ' Si les clés existent déjà, on remplace la valeur associée
                    Err.Raise 1014, , "Keys " & Join(keys, ", ") & " already exist in the dictionary."
                    Exit Sub
                Else
                    ' Ajoute les cl?s et la valeur si les cl?s n'existent pas d?j?
                    For k = LBound(keys) To UBound(keys)
                        nKeys(i).Add keys(i)
                    Next k
                    values.Add value
                End If
            Next i
        ' Ajoute les clés et la valeur si les clés n'existent pas déjà
        For i = LBound(keys) To UBound(keys)
            nKeys(i).Add keys(i)
        Next i
        values.Add value
    End If
End Sub




' Vérifie si une clé existe dans le dictionnaire
' keys : Un tableau de clés à vérifier
' Retourne : True si la clé existe, False sinon
Public Function Exists(keys() As Variant) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Parcourt les collections de clés pour vérifier l'existence de la clé
        For i = 1 To nKeys(1).count
            found = True
            For j = LBound(keys) To UBound(keys)
                If nKeys(j)(i) <> keys(j) Then
                    found = False
                    Exit For
                End If
            Next j
            If found Then
                Exists = True
                Exit Function
            End If
        Next i
    End If
    ' Si la clé n'est pas trouvée, retourne False
    Exists = False
End Function

' Vérifie si une clé est dans une collection
' col : La collection à vérifier
' key : La clé à vérifier
' Retourne : True si la clé est dans la collection, False sinon
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

' Supprime une paire cl?-valeur du dictionnaire
' keys : Un tableau de cl?s de la paire ? supprimer
Public Function Remove(keys() As Variant)
    Dim i, j As Integer
    Dim hasAllKey As Boolean
    Dim Dimen As Collection
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(keys) - LBound(keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Parcourt les collections afin de tester si la valeur est égale à la clé donné en parmaètre
            'Parcours toutes les clés
            For i = 1 To nKeys(1).count
            hasAllKey = True
            
            'Parcours toutes les dimensions
                For j = 1 To Me.dimension
                    hasAllKey = (IIf(nKeys(j)(i) = keys(j), 1, 0)) * hasAllKey
                    If hasAllKey = False Then
                        Exit For
                    End If
                Next j
                'Si on a trouvé toutes les clés alors on retourne la valeur
                If hasAllKey = True Then
                    ' Retourne la valeur
                    values.Remove (i)
                    For Each Dimen In nKeys
                        Dimen.Remove (i)
                    Next
                Exit Function
            End If
            Next i

        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        If hasAllKey = False Then
            Err.Raise 1013, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
        End If
    End If
    On Error GoTo 0
End Function




' Obtient la valeur associée à une paire de clés
' keys : Un tableau de clés dont la valeur doit être obtenue
' Retourne : La valeur associée aux clés
Public Function item(keys() As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim hasAllKey As Boolean
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(keys) - LBound(keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Parcourt les collections afin de tester si la valeur est égale à la clé donné en parmaètre
            'Parcours toutes les clés
            For i = 1 To nKeys(1).count
            hasAllKey = True
            
            'Parcours toutes les dimensions
                For j = 1 To Me.dimension
                    hasAllKey = (IIf(nKeys(j)(i) = keys(j), 1, 0)) * hasAllKey
                    If hasAllKey = False Then
                        Exit For
                    End If
                Next j
                'Si on a trouvé toutes les clés alors on retourne la valeur
                If hasAllKey = True Then
                    ' Retourne la valeur
                    item = values(i)
                Exit Function
            End If
            Next i

        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        If hasAllKey = False Then
            Err.Raise 1013, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
        End If
    End If
    On Error GoTo 0
End Function

' Obtient toutes les valeurs associées à une clé donnée
' keys : Un tableau de clés dont les valeurs doivent être obtenues
' keepKey : Un paramètre optionnel pour déterminer si la clé doit être conservée dans le nouveau dictionnaire
' Retourne : Un objet DictionnaryND contenant les valeurs associées aux clés
Public Function items(keys() As Variant, Optional keepKey As Boolean = True) As DictionnaryND
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim hasKey As Boolean
    Dim newKeys() As Variant
    Dim newDict As DictionnaryND
    Set newDict = New DictionnaryND
    
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dictionary has not been initialize."
    Else
        ' Parcourt les collections afin de tester si la valeur est égale à la clé donné en parmaètre
            'Parcours toutes les clés
            For i = 1 To nKeys(1).count
            hasAllKey = True
            
            'Parcours toutes les dimensions
                For j = 1 To Me.dimension
                    'Test si la clé fourni est vide ou Nothing
                    If keys(j) <> Nothing Or keys(j) <> "" Then
                        'Si la clé n'est pas vide ou empty
                        'On calcul si la clé match parmis les nKeys on obtiens "1" sinon "0".
                        'Le produit de ces valeurs permets de faire un "ET" logique et de savoir si toutes les clés match
                        hasAllKey = (IIf(nKeys(j)(i) = keys(j), 1, 0)) * hasAllKey
                        If hasAllKey = False Then
                            Exit For
                        End If
                    End If
                Next j
                'Si on a trouvé toutes les clés alors ajoute la valeur dans le dictionnaire
                If hasAllKey = True Then
                    ' Cr?e un nouveau tableau de cl?s pour le nouveau dictionnaire
                    ReDim newKeys(1 To Me.dimension)
                    For l = 1 To Me.dimension
                        If IsEmpty(keys(l)) Then
                            newKeys(l) = nKeys(l)(i)
                        Else
                            newKeys(l) = keys(l)
                        End If
                    Next l
                    ' Ajoute la valeur au nouveau dictionnaire
                    newDict.Add values(i), newKeys
                Exit Function
            End If
            Next i
        
        ' Retourne le nouveau dictionnaire
        Set items = newDict
    End If
End Function




' Crée un tableau 2D à partir du dictionnaire
' Retourne : Un tableau 2D où chaque ligne est une paire clé-valeur du dictionnaire
Public Function toArray() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim arr() As Variant
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        Err.Raise 1011, , "Dimension of the dictionary has not been specified."
    Else
        ' Redimensionne le tableau pour contenir toutes les paires clé-valeur
        ReDim arr(1 To nKeys(1).count, 1 To dimension + 1)
        ' Remplit le tableau avec les clés et les valeurs
        For i = 1 To nKeys(1).count
            For j = 1 To dimension
                arr(i, j) = nKeys(j)(i)
            Next j
            arr(i, dimension + 1) = values(i)
        Next i
    End If
    ' Retourne le tableau
    toArray = arr
End Function

' Initialise le dictionnaire à partir d'un tableau 2D
' arr : Un tableau 2D où chaque ligne est une paire clé-valeur
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

        Set values = New Collection
        ' Remplit le dictionnaire avec les clés et les valeurs
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2) - 1
                nKeys(j).Add arr(i, j)
            Next j
            values.Add arr(i, UBound(arr, 2))
        Next i
    End If
End Function





Sub Test()
    Dim dict As New DictionnaryND
    Dim keys(1 To 2) As Variant
    Dim value As Variant
    Dim i, j As Integer
    
    ' Initialise le dictionnaire avec 2 dimensions
    dict.dimension = 2
    
    ' Ajoute des paires clé-valeur au dictionnaire
    For i = 1 To 10
        For j = 1 To 10
            keys(1) = "Key1_" & i & j
            keys(2) = "Key2_" & i & j
            value = "Value_" & i & j
            dict.Add value, keys
        Next j
    Next i
    
    'Initialisation des tableaux de sorti
    Dim arr() As Variant
    Dim arr2() As Variant
    Dim arr3() As Variant
    
    arr = dict.toArray
    
    Dim dictFromArray As New DictionnaryND
    dictFromArray.FromArray arr

    arr2 = dictFromArray.toArray
    
    
    Dim result As String
    result = dictFromArray.item(keys)
    If Not result = value Then
        Debug.Print "ERROR - Got the item :" & result & " for the keys {" & keys(1) & "," & keys(2) & "} should have gotten " & value
    Else
        Debug.Print "VALID - Got the item :" & result & " for the keys {Key1_10, Key2_10} - VALID"
    End If
    
    ' Supprime des paires clé-valeur du dictionnaire
    For i = 1 To 5
        keys(1) = "Key1_" & i & "1"
        keys(2) = "Key2_" & i & "1"
        dict.Remove keys
    Next i
    
    ' Vérifie à nouveau l'existence des clés et obtient les valeurs
    For i = 1 To 10
        keys(1) = "Key1_" & i & "1"
        keys(2) = "Key2_" & i & "1"
        If dict.Exists(keys) Then
            Debug.Print "Keys: " & Join(keys, ", ") & ", Value: " & dict.item(keys)
        Else
            Debug.Print "Keys: " & Join(keys, ", ") & " not found in the dictionary."
        End If
    Next i
    
    Dim slice As DictionnaryND
    keys(2) = Empty
    Set slice = dictFromArray.items(keys) ' doesn't return the dict
    arr3 = slice.toArray
    
    MsgBox "End of test"
    
End Sub

