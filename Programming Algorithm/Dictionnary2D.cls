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

Private keys1 As Collection ' Premi�re collection pour stocker les cl�s
Private keys2 As Collection ' Deuxi�me collection pour stocker les cl�s
Private values As Collection ' Collection pour stocker les valeurs

' Initialise les collections lors de la cr�ation de l'objet Dictionary2D
Private Sub Class_Initialize()
    Set keys1 = New Collection
    Set keys2 = New Collection
    Set values = New Collection
End Sub

' Ajoute une nouvelle paire cl�-valeur au dictionnaire
' key1 : La premi�re cl� � ajouter
' key2 : La deuxi�me cl� � ajouter
' value : La valeur � ajouter
Public Sub Add(key1 As Variant, key2 As Variant, value As Variant)
    On Error Resume Next
    ' V�rifie si les cl�s existent d�j�
    If IsInCollection(keys1, key1) And IsInCollection(keys2, key2) Then
        Err.Raise 1010, , "Keys " & key1 & " and " & key2 & " already exist in the dictionary."
    Else
        keys1.Add key1
        keys2.Add key2
        values.Add value
    End If
    On Error GoTo 0
End Sub

' V�rifie si une paire de cl�s existe dans le dictionnaire
' key1 : La premi�re cl� � v�rifier
' key2 : La deuxi�me cl� � v�rifier
' Retourne : True si les cl�s existent, False sinon
Public Function Exists(key1 As Variant, key2 As Variant) As Boolean
    On Error Resume Next
    Exists = Not IsEmpty(keys1(key1)) And Not IsEmpty(keys2(key2))
    On Error GoTo 0
End Function

' Supprime une paire cl�-valeur du dictionnaire
' key1 : La premi�re cl� de la paire � supprimer
' key2 : La deuxi�me cl� de la paire � supprimer
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
    ' Si les cl�s ne sont pas trouv�es, l�ve une erreur
    If Not found Then
        Err.Raise 1011, , "Keys " & key1 & " and " & key2 & " not found in the dictionary."
    End If
End Function

' Obtient la valeur associ�e � une paire de cl�s
' key1 : La premi�re cl� dont la valeur doit �tre obtenue
' key2 : La deuxi�me cl� dont la valeur doit �tre obtenue
' Retourne : La valeur associ�e aux cl�s
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
    ' Si les cl�s ne sont pas trouv�es, l�ve une erreur
    If Not found Then
        Err.Raise 1012, , "Keys " & key1 & " and " & key2 & " not found in the dictionary."
    End If
End Function

' Obtient toutes les valeurs associ�es � une cl� donn�e
' key : La cl� dont les valeurs doivent �tre obtenues
' Retourne : Un objet Dictionnary contenant les valeurs associ�es � la cl�
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
    ' Si la cl� n'est pas trouv�e, l�ve une erreur
    If Not found Then
        Err.Raise 1013, , "Key " & key & " not found in the dictionary."
    Else
        Set items = results
    End If
End Function


' V�rifie si une cl� est dans une collection
' col : La collection � v�rifier
' key : La cl� � v�rifier
' Retourne : True si la cl� est dans la collection, False sinon
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
