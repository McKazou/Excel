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

' Ajoute une nouvelle paire clé-valeur au dictionnaire
' value : La valeur à ajouter
' keys : Un tableau de clés à ajouter
' Comportement : Dans le cas ou la clé est déjà présente dans le dictionnaire, pour le moment une erreur est levé, il est possible de "remplacer" la valeur à la place
Public Sub Add(value As Variant, keys() As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        ' Initialise le dictionnaire avec les dimensions de cette première sé
        Me.dimension = UBound(keys) - LBound(keys) + 1
    End If
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(keys) - LBound(keys) + 1 <> dimension Then
        Err.Raise 1013, , "Number of keys does not match the number of dimensions."
    Else
        ' Vérifie si les clés existent déjà
        For i = 1 To nKeys(1).count
            found = True
            For j = LBound(keys) To UBound(keys)
                If nKeys(j + 1)(i) <> keys(j) Then
                    found = False
                    Exit For
                End If
            Next j
            If found Then
                Err.Raise 1014, , "Keys " & Join(keys, ", ") & " already exist in the dictionary."
                Exit Sub
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
        Err.Raise 1012, , "Dimension of the dictionary has not been specified."
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


Public Function Remove(keys() As Variant)
    Dim dimension As Collection
    Dim keyInCollection As Variant
    Dim i As Integer
    Dim found As Boolean
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dimension of the dictionary has not been specified."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(keys) - LBound(keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Parcourt les collections de clés pour trouver et supprimer la paire clé-valeur
            i = 1
            For Each dimension In nKeys
                found = True
                For j = LBound(keys) To UBound(keys)
                    If dimension(j) <> keys(j) Then
                        found = False
                        Exit For
                    End If
                Next j
                If found Then
                    ' Supprime la paire clé-valeur
                    dimension.Remove i
                    values.Remove i
                    Exit Function
                End If
                i = i + 1
            Next dimension
        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        If Not found Then
            Err.Raise 1012, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
        End If
    End If
    On Error GoTo 0
End Function




' Obtient la valeur associée à une paire de clés
' keys : Un tableau de clés dont la valeur doit être obtenue
' Retourne : La valeur associée aux clés
Public Function item(keys() As Variant) As Variant
    Dim dimension As Collection
    Dim keyInCollection As Variant
    Dim key As Variant
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If Me.dimension = 0 Then
        Err.Raise 1012, , "Dimension of the dictionary has not been specified."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(keys) - LBound(keys) + 1 <> Me.dimension Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Parcourt les collections de clés pour trouver la valeur
            For i = 1 To nKeys(1).count
                found = True
                j = 1
                For Each dimension In nKeys
                    For Each keyInCollection In dimension
                        If keyInCollection = keys(j) Then
                            Exit For
                        End If
                    Next keyInCollection
                    j = j + 1
                Next dimension
                If found Then
                    ' Retourne la valeur
                    item = values(i)
                    Exit Function
                End If
            Next i
        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        If Not found Then
            Err.Raise 1013, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
        End If
    End If
    On Error GoTo 0
End Function




' Obtient toutes les valeurs associées à une clé donnée
' keys : Un tableau de clés dont les valeurs doivent être obtenues
' Retourne : Un objet DictionnaryND contenant les valeurs associées aux clés
Public Function items(keys() As Variant) As DictionnaryND
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim results As New DictionnaryND
    On Error Resume Next
    ' Vérifie si la dimension du dictionnaire a été spécifiée
    If dimension = 0 Then
        Err.Raise 1012, , "Dimension of the dictionary has not been specified."
    Else
        ' Vérifie si le nombre de clés correspond au nombre de dimensions
        If UBound(keys) - LBound(keys) + 1 <> n Then
            Err.Raise 1013, , "Number of keys does not match the number of dimensions."
        Else
            ' Parcourt les collections de clés pour trouver les valeurs
            For i = 1 To nKeys(1).count
                found = True
                For j = LBound(keys) To UBound(keys) - 1
                    If nKeys(j + 1)(i) <> keys(j) Then
                        found = False
                        Exit For
                    End If
                Next j
                If found And nKeys(UBound(keys))(i) = keys(UBound(keys)) Then
                    ' Ajoute la valeur à l'objet DictionnaryND de résultats
                    results.Add values(i)
                End If
            Next i
        End If
        ' Si les clés ne sont pas trouvées, lève une erreur
        If Not found Then
            Err.Raise 1014, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
        Else
            Set items = results
        End If
    End If
    On Error GoTo 0
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



Sub Test()
    Dim dict As New DictionnaryND
    Dim keys(1 To 2) As Variant
    Dim value As Variant
    Dim i As Integer
    
    ' Initialise le dictionnaire avec 2 dimensions
    dict.dimension = 2
    
    ' Ajoute des paires clé-valeur au dictionnaire
    For i = 1 To 10
        keys(1) = "Key1_" & i
        keys(2) = "Key2_" & i
        value = "Value_" & i
        dict.Add value, keys
    Next i
    Dim arr() As Variant
    arr = dict.toArray
    
    ' Vérifie l'existence des clés et obtient les valeurs
    For i = 1 To 10
        keys(1) = "Key1_" & i
        keys(2) = "Key2_" & i
        If dict.Exists(keys) Then
            Debug.Print "Keys: " & Join(keys, ", ") & ", Value: " & dict.item(keys)
        Else
            Debug.Print "Keys: " & Join(keys, ", ") & " not found in the dictionary."
        End If
    Next i
    
    ' Supprime des paires clé-valeur du dictionnaire
    For i = 1 To 5
        keys(1) = "Key1_" & i
        keys(2) = "Key2_" & i
        dict.Remove keys
    Next i
    
    ' Vérifie à nouveau l'existence des clés et obtient les valeurs
    For i = 1 To 10
        keys(1) = "Key1_" & i
        keys(2) = "Key2_" & i
        If dict.Exists(keys) Then
            Debug.Print "Keys: " & Join(keys, ", ") & ", Value: " & dict.item(keys)
        Else
            Debug.Print "Keys: " & Join(keys, ", ") & " not found in the dictionary."
        End If
    Next i
End Sub

