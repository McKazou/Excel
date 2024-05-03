Attribute VB_Name = "RandomGenerator"
Function GenerateRandomText(length As Integer, Optional IncludeNumbers As Boolean = True, Optional IncludeSymbols As Boolean = True, Optional IncludeSpaces As Boolean = True) As String
    Dim characters As String
    Dim i As Integer
    Dim randomCharacter As String
    
    ' D�finir les caract�res qui peuvent �tre utilis�s dans le texte al�atoire
    characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If IncludeNumbers Then characters = characters & "0123456789"
    If IncludeSymbols Then characters = characters & "!@#$%^&*()_-+={}[]|\:;'""<>,.?/"
    If IncludeSpaces Then characters = characters & " "
    
    ' Initialiser la cha�ne de texte al�atoire
    GenerateRandomText = ""
    
    ' G�n�rer le texte al�atoire
    For i = 1 To length
        ' S�lectionner un caract�re al�atoire
        randomCharacter = Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
        ' Ajouter le caract�re al�atoire au texte al�atoire
        GenerateRandomText = GenerateRandomText & randomCharacter
    Next i
End Function


