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
Private n As Integer ' Le nombre de dimensions du dictionnaire

' Initialise les collections lors de la création de l'objet DictionaryND
' n : Le nombre de dimensions du dictionnaire
Public Sub Class_Initialize(n As Integer)
    Dim i As Integer
    Set nKeys = New Collection
    Set values = New Collection
    n = n
    ' Crée une nouvelle collection de clés pour chaque dimension
    For i = 1 To n
        nKeys.Add New Collection
    Next i
End Sub

' Ajoute une nouvelle paire clé-valeur au dictionnaire
' value : La valeur à ajouter
' keys : Un tableau de clés à ajouter
Public Sub Add(value As Variant, keys() As Variant)
    Dim i As Integer
    On Error Resume Next
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(keys) - LBound(keys) + 1 <> n Then
        Err.Raise 1010, , "Number of keys does not match the number of dimensions."
    Else
        ' Vérifie si les clés existent déjà et les ajoute si ce n'est pas le cas
        For i = LBound(keys) To UBound(keys)
            If Not IsInCollection(nKeys(i + 1), keys(i)) Then
                nKeys(i + 1).Add keys(i)
            End If
        Next i
        values.Add value
    End If
    On Error GoTo 0
End Sub

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


' Supprime une paire clé-valeur du dictionnaire
' keys : Un tableau de clés de la paire à supprimer
Public Function Remove(keys() As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    On Error Resume Next
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(keys) - LBound(keys) + 1 <> n Then
        Err.Raise 1011, , "Number of keys does not match the number of dimensions."
    Else
        ' Parcourt les collections de clés pour trouver et supprimer la paire clé-valeur
        For i = 1 To nKeys(1).count
            found = True
            For j = LBound(keys) To UBound(keys)
                If nKeys(j + 1)(i) <> keys(j) Then
                    found = False
                    Exit For
                End If
            Next j
            If found Then
                ' Supprime la paire clé-valeur
                For j = LBound(keys) To UBound(keys)
                    nKeys(j + 1).Remove i
                Next j
                values.Remove i
                Exit Function
            End If
        Next i
    End If
    ' Si les clés ne sont pas trouvées, lève une erreur
    If Not found Then
        Err.Raise 1012, , "Keys " & Join(keys, ", ") & " not found in the dictionary."
    End If
    On Error GoTo 0
End Function


' Obtient la valeur associée à une paire de clés
' keys : Un tableau de clés dont la valeur doit être obtenue
' Retourne : La valeur associée aux clés
Public Function item(keys() As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    On Error Resume Next
    ' Vérifie si le nombre de clés correspond au nombre de dimensions
    If UBound(keys) - LBound(keys) + 1 <> n Then
        Err.Raise 1012, , "Number of keys does not match the number of dimensions."
    Else
        ' Parcourt les collections de clés pour trouver la valeur
        For i = 1 To nKeys(1).count
            found = True
            For j = LBound(keys) To UBound(keys)
                If nKeys(j + 1)(i) <> keys(j) Then
                    found = False
                    Exit For
                End If
            Next j
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
    On Error GoTo 0
End Function


Sub Test()
    Dim dict As New DictionnaryND
    Dim keys(1 To 2) As Variant
    Dim value As Variant
    Dim i As Integer
    
    ' Initialise le dictionnaire avec 2 dimensions
    dict.Class_Initialize 2
    
    ' Ajoute des paires clé-valeur au dictionnaire
    For i = 1 To 10
        keys(1) = "Key1_" & i
        keys(2) = "Key2_" & i
        value = "Value_" & i
        dict.Add value, keys
    Next i
    
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

