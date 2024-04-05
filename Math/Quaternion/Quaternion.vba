' Dans le module de classe Quaternion
Public a As Double ' Composante réelle
Public b As Double ' Composante imaginaire i
Public c As Double ' Composante imaginaire j
Public d As Double ' Composante imaginaire k

' Méthode pour calculer la norme du Quaternion
Public Function Norme() As Double
    Norme = Sqr(a ^ 2 + b ^ 2 + c ^ 2 + d ^ 2)
End Function

' Méthode pour additionner le Quaternion avec un autre Quaternion
Public Function Addition(q2 As Quaternion) As Quaternion
    Dim q3 As New Quaternion
    q3.a = Me.a + q2.a
    q3.b = Me.b + q2.b
    q3.c = Me.c + q2.c
    q3.d = Me.d + q2.d
    Set Addition = q3
End Function

' Méthode pour multiplier le Quaternion par un scalaire
Public Function Multiplication(k As Double) As Quaternion
    Dim q4 As New Quaternion
    q4.a = Me.a * k
    q4.b = Me.b * k
    q4.c = Me.c * k
    q4.d = Me.d * k
    Set Multiplication = q4
End Function

' Méthode pour multiplier le Quaternion avec un autre Quaternion
Public Function Multiplication(q2 As Quaternion) As Quaternion
    Dim q3 As New Quaternion
    q3.a = Me.a * q2.a - Me.b * q2.b - Me.c * q2.c - Me.d * q2.d ' Composante réelle
    q3.b = Me.a * q2.b + Me.b * q2.a + Me.c * q2.d - Me.d * q2.c ' Composante imaginaire i
    q3.c = Me.a * q2.c - Me.b * q2.d + Me.c * q2.a + Me.d * q2.b ' Composante imaginaire j
    q3.d = Me.a * q2.d + Me.b * q2.c - Me.c * q2.b + Me.d * q2.a ' Composante imaginaire k
    Set Multiplication = q3
End Function

' Dans un module standard
Sub Test()
    Dim q1 As New Quaternion ' Création d'un objet q1 de type Quaternion
    Dim q2 As New Quaternion ' Création d'un objet q2 de type Quaternion
    Dim q3 As Quaternion ' Déclaration d'un objet q3 de type Quaternion
    Dim q4 As Quaternion ' Déclaration d'un objet q4 de type Quaternion
   
    ' Affectation des valeurs aux composantes des Quaternions
    q1.a = 1
    q1.b = 2
    q1.c = 3
    q1.d = 4
    q2.a = 5
    q2.b = 6
    q2.c = 7
    q2.d = 8
   
    ' Utilisation des méthodes de la classe Quaternion
    Debug.Print q1.Norme ' Affiche la norme de q1
    Set q3 = q1.Addition(q2) ' Affecte à q3 la somme de q1 et q2
    Debug.Print q3.a, q3.b, q3.c, q3.d ' Affiche les composantes de q3
    Set q4 = q1.Multiplication(2) ' Affecte à q4 le produit de q1 par 2
    Debug.Print q4.a, q4.b, q4.c, q4.d ' Affiche les composantes de q4
End Sub
