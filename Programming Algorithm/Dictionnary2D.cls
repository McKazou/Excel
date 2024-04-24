VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionnary2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private keys1 As Collection ' Première collection pour stocker les clés
Private keys2 As Collection ' Deuxième collection pour stocker les clés
Private values As Collection ' Collection pour stocker les valeurs

' Initialise les collections lors de la création de l'objet Dictionary2D
Private Sub Class_Initialize()
    Set keys1 = New Collection
    Set keys2 = New Collection
    Set values = New Collection
End Sub

' Ajoute une nouvelle paire clé-valeur au dictionnaire
' key1 : La première clé à ajouter
' key2 : La deuxième clé à ajouter
' value : La valeur à ajouter
Public Sub Add(key1 As Variant, key2 As Variant, value As Variant)
    On Error Resume Next
    ' Vérifie si les clés existent déjà
    If IsInCollection(keys1, key1) And IsInCollection(keys2, key2) Then
        Err.Raise 1010, , "Keys " & key1 & " and " & key2 & " already exist in the dictionary."
    Else
        keys1.Add key1
        keys2.Add key2
        values.Add value
    End If
    On Error GoTo 0
End Sub

' Vérifie si une paire de clés existe dans le dictionnaire
' key1 : La première clé à vérifier
' key2 : La deuxième clé à vérifier
' Retourne : True si les clés existent, False sinon
Public Function Exists(key1 As Variant, key2 As Variant) As Boolean
    On Error Resume Next
    Exists = Not IsEmpty(keys1(key1)) And Not IsEmpty(keys2(key2))
    On Error GoTo 0
End Function

' Supprime une paire clé-valeur du dictionnaire
' key1 : La première clé de la paire à supprimer
' key2 : La deuxième clé de la paire à supprimer
Public Function Remove(key1 As Variant, key2 As Variant)
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 1 To keys1.count
        If keys1(i) = key1 And keys2(i) = key2 Then
            keys1.Remove i
            keys2.Remove i
            values.Remove i
            found = True
            Exit Function
        End If
    Next i
    ' Si les clés ne sont pas trouvées, lève une erreur
    If Not found Then
        Err.Raise 1011, , "Keys " & key1 & " and " & key2 & " not found in the dictionary."
    End If
End Function

' Obtient la valeur associée à une paire de clés
' key1 : La première clé dont la valeur doit être obtenue
' key2 : La deuxième clé dont la valeur doit être obtenue
' Retourne : La valeur associée aux clés
Public Function item(key1 As Variant, key2 As Variant) As Variant
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 1 To keys1.count
        If keys1(i) = key1 And keys2(i) = key2 Then
            item = values(i)
            found = True
            Exit Function
        End If
    Next i
    ' Si les clés ne sont pas trouvées, lève une erreur
    If Not found Then
        Err.Raise 1012, , "Keys " & key1 & " and " & key2 & " not found in the dictionary."
    End If
End Function

' Obtient toutes les valeurs associées à une clé donnée
' key : La clé dont les valeurs doivent être obtenues
' Retourne : Un objet Dictionnary contenant les valeurs associées à la clé
Public Function items(key As Variant) As Dictionnary
    Dim i As Long
    Dim found As Boolean
    Dim results As New Dictionnary
    found = False
    For i = 1 To keys1.count
        If keys1(i) = key Then
            results.Add keys2(i), values(i)
            found = True
        End If
    Next i
    ' Si la clé n'est pas trouvée, lève une erreur
    If Not found Then
        Err.Raise 1013, , "Key " & key & " not found in the dictionary."
    Else
        Set items = results
    End If
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

Sub Test()
    Dim dict As New Dictionnary2D
    Dim key1 As Variant, key2 As Variant
    Dim value As Variant
    Dim i As Integer
    
    ' Ajoute des paires clé-valeur au dictionnaire
    For i = 1 To 10
        key1 = "Key1_" & i
        key2 = "Key2_" & i
        value = "Value_" & i
        dict.Add key1, key2, value
    Next i
    
    ' Vérifie l'existence des clés et obtient les valeurs
    For i = 1 To 10
        key1 = "Key1_" & i
        key2 = "Key2_" & i
        If dict.Exists(key1, key2) Then
            Debug.Print "Keys: " & key1 & ", " & key2 & ", Value: " & dict.item(key1, key2)
        Else
            Debug.Print "Keys: " & key1 & ", " & key2 & " not found in the dictionary."
        End If
    Next i
    
    ' Supprime des paires clé-valeur du dictionnaire
    For i = 1 To 5
        key1 = "Key1_" & i
        key2 = "Key2_" & i
        dict.Remove key1, key2
    Next i
    
    ' Vérifie à nouveau l'existence des clés et obtient les valeurs
    For i = 1 To 10
        key1 = "Key1_" & i
        key2 = "Key2_" & i
        If dict.Exists(key1, key2) Then
            Debug.Print "Keys: " & key1 & ", " & key2 & ", Value: " & dict.item(key1, key2)
        Else
            Debug.Print "Keys: " & key1 & ", " & key2 & " not found in the dictionary."
        End If
    Next i
End Sub

