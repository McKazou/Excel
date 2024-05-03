Attribute VB_Name = "RandomGenerator"
Function GenerateRandomText(length As Integer, Optional IncludeNumbers As Boolean = True, Optional IncludeSymbols As Boolean = True, Optional IncludeSpaces As Boolean = True) As String
    Dim characters As String
    Dim i As Integer
    Dim randomCharacter As String
    
    ' Définir les caractères qui peuvent être utilisés dans le texte aléatoire
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If IncludeNumbers Then characters = characters & "0123456789"
    If IncludeSymbols Then characters = characters & "!@#$%^&*()_-+={}[]|\:;'""<>,.?/"
    If IncludeSpaces Then characters = characters & " "
    
    ' Initialiser la chaîne de texte aléatoire
    GenerateRandomText = ""
    
    ' Générer le texte aléatoire
    For i = 1 To length
        ' Sélectionner un caractère aléatoire
        randomCharacter = Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
        ' Ajouter le caractère aléatoire au texte aléatoire
        GenerateRandomText = GenerateRandomText & randomCharacter
    Next i
End Function


