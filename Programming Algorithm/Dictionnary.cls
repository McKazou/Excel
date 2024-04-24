VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionnary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private keys As Collection ' Collection pour stocker les clés
Private values As Collection ' Collection pour stocker les valeurs

' Initialise les collections lors de la création de l'objet Dictionary
Private Sub Class_Initialize()
    Set keys = New Collection
    Set values = New Collection
End Sub

' Ajoute une nouvelle paire clé-valeur au dictionnaire
' key : La clé à ajouter
' value : La valeur à ajouter
Public Sub Add(key As Variant, value As Variant)
    On Error Resume Next
    ' Vérifie si la clé existe déjà
    If IsInCollection(keys, key) Then
        Err.Raise 1010, , "Key " & key & " already exists in the dictionary."
    Else
        keys.Add key
        values.Add value
    End If
    On Error GoTo 0
End Sub

' Vérifie si une clé existe dans le dictionnaire
' key : La clé à vérifier
' Retourne : True si la clé existe, False sinon
Public Function Exists(key As Variant) As Boolean
    On Error Resume Next
    Exists = Not IsEmpty(keys(key))
    On Error GoTo 0
End Function

' Supprime une paire clé-valeur du dictionnaire
' key : La clé de la paire à supprimer
Public Function Remove(key As Variant)
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 1 To keys.count
        If keys(i) = key Then
            keys.Remove i
            values.Remove i
            found = True
            Exit Function
        End If
    Next i
    ' Si la clé n'est pas trouvée, lève une erreur
    If Not found Then
        Err.Raise 1011, , "Key " & key & " not found in the dictionary."
    End If
End Function


' Obtient la valeur associée à une clé
' key : La clé dont la valeur doit être obtenue
' Retourne : La valeur associée à la clé
Public Function item(key As Variant) As Variant
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 1 To keys.count
        If keys(i) = key Then
            item = values(i)
            found = True
            Exit Function
        End If
    Next i
    ' Si la clé n'est pas trouvée, lève une erreur
    If Not found Then
        Err.Raise 1012, , "Key " & key & " not found in the dictionary."
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

