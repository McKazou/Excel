VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quaternion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Last update 240413 - 01h59
'Add Reference to Regex
' Dans le module de classe Quaternion
Public a As Double ' Composante rÃ©elle
Public b As Double ' Composante imaginaire i
Public c As Double ' Composante imaginaire j
Public d As Double ' Composante imaginaire k

' Constructeur de la classe
Public Sub Class_Initialize()
    On Error GoTo 0
    Dim xlma As New XlmaManager
    xlma.refreshModuleByName "Math_Quaternion", ThisWorkbook.VBProject
End Sub

Public Sub fromString(s As String)
    Dim a As Double, b As Double, c As Double, d As Double
    Dim regex As Object, matches As Object

    ' Créer une nouvelle instance de l'objet RegExp
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "([-+]?\d*\.?\d+)?([ijk])?"
    End With

    ' Utiliser l'objet RegExp pour trouver des correspondances dans la chaîne
    Set matches = regex.Execute(s)

    ' Initialiser les valeurs a, b, c et d
    a = 0
    b = 0
    c = 0
    d = 0

    ' Parcourir les correspondances et affecter les valeurs à a, b, c et d
    Dim match As Object
    For Each match In matches
        Select Case match.SubMatches(1)
            Case "i"
                If match.SubMatches(0) = "" Then
                    b = 1
                Else
                    b = CDbl(match.SubMatches(0))
                End If
            Case "j"
                If match.SubMatches(0) = "" Then
                    c = 1
                Else
                    c = CDbl(match.SubMatches(0))
                End If
            Case "k"
                If match.SubMatches(0) = "" Then
                    d = 1
                Else
                    d = CDbl(match.SubMatches(0))
                End If
            Case Else
                If match.SubMatches(0) <> "" Then
                    a = CDbl(match.SubMatches(0))
                End If
        End Select
    Next match

    ' Afficher les valeurs a, b, c et d
    Me.a = a
    Me.b = b
    Me.c = c
    Me.d = d
End Sub



Public Function toString(Optional numDecimals As Integer = -1) As String
    Dim str As String
    str = ""
    If Me.a <> 0 Then
        If numDecimals >= 0 Then
            str = str & Round(Me.a, numDecimals)
        Else
            str = str & Me.a
        End If
    End If
    If Me.b <> 0 Then
        If Me.b = 1 Then
            str = str & "+i"
        ElseIf Me.b = -1 Then
            str = str & "-i"
        Else
            Dim roundedB As Double
            If numDecimals >= 0 Then
                roundedB = Round(Me.b, numDecimals)
                If roundedB <> 0 Then
                    str = str & "+" & roundedB & "i"
                End If
            Else
                str = str & "+" & Me.b & "i"
            End If
        End If
    End If
    If Me.c <> 0 Then
        If Me.c = 1 Then
            str = str & "+j"
        ElseIf Me.c = -1 Then
            str = str & "-j"
        Else
            Dim roundedC As Double
            If numDecimals >= 0 Then
                roundedC = Round(Me.c, numDecimals)
                If roundedC <> 0 Then
                    str = str & "+" & roundedC & "j"
                End If
            Else
                str = str & "+" & Me.c & "j"
            End If
        End If
    End If
    If Me.d <> 0 Then
        If Me.d = 1 Then
            str = str & "+k"
        ElseIf Me.d = -1 Then
            str = str & "-k"
        Else
            Dim roundedD As Double
            If numDecimals >= 0 Then
                roundedD = Round(Me.d, numDecimals)
                If roundedD <> 0 Then
                    str = str & "+" & roundedD & "k"
                End If
            Else
                str = str & "+" & Me.d & "k"
            End If
        End If
    End If
    ' Supprimer le premier caractère s'il s'agit d'un '+'
    If Left(str, 1) = "+" Then
        str = Mid(str, 2)
    End If
    toString = str
End Function







' Convertir un Quaternion en tableau
Public Function toArray() As Variant

    toArray = Array(Me.a, Me.b, Me.c, Me.d)
End Function

'Converti une range en array
Public Sub fromRange(rng As Range)
    Dim rangeAsArray() As Variant
    rangeAsArray = rng.Value2
    
    Me.fromArray rangeAsArray
End Sub

' Convertir un tableau en Quaternion
Public Sub fromArray(arr() As Variant)
    Dim numElements As Long
    numElements = (UBound(arr, 1) - LBound(arr, 1) + 1) * (UBound(arr, 2) - LBound(arr, 2) + 1)

    If numElements = 1 Then
        Me.fromString (CStr(arr(1, 1)))
    ElseIf numElements = 3 Then
        ' Si arr est un tableau 2D (1 to 1, 0 to 2)
        Me.a = 0 ' Composante réelle
        Me.b = CDbl(arr(1, 1)) ' Composante imaginaire i
        Me.c = CDbl(arr(1, 2)) ' Composante imaginaire j
        Me.d = CDbl(arr(1, 3)) ' Composante imaginaire k
    ElseIf numElements = 4 Then
        ' Si arr est un tableau 2D (1 to 1, 0 to 3)
        Me.a = CDbl(arr(1, 1)) ' Composante réelle
        Me.b = CDbl(arr(1, 2)) ' Composante imaginaire i
        Me.c = CDbl(arr(1, 3)) ' Composante imaginaire j
        Me.d = CDbl(arr(1, 4)) ' Composante imaginaire k
    Else
        Err.Raise 4, , "Nombre d'éléments fourni dans l'array non géré : " & numElements
    End If
End Sub



' Convertit les angles d'Euler en un quaternion
Sub fromEulerAngle(rotX As Double, rotY As Double, rotZ As Double, isRad As Boolean, Optional convention As String = "ZYX")

    Dim a As Double, b As Double, c As Double

    ' Convertir les angles en radians si nécessaire
    If Not isRad Then
        rotX = WorksheetFunction.Radians(rotX)
        rotY = WorksheetFunction.Radians(rotY)
        rotZ = WorksheetFunction.Radians(rotZ)
    End If

    a = rotX / 2
    b = rotY / 2
    c = rotZ / 2

    Select Case convention
        Case "ZYX"
            Me.a = Cos(a) * Cos(b) * Cos(c) + Sin(a) * Sin(b) * Sin(c)
            Me.b = Sin(a) * Cos(b) * Cos(c) - Cos(a) * Sin(b) * Sin(c)
            Me.c = Cos(a) * Sin(b) * Cos(c) + Sin(a) * Cos(b) * Sin(c)
            Me.d = Cos(a) * Cos(b) * Sin(c) - Sin(a) * Sin(b) * Cos(c)
        Case "XYZ"
            Me.a = Cos(a) * Cos(b) * Cos(c) - Sin(a) * Sin(b) * Sin(c)
            Me.b = Sin(a) * Cos(b) * Cos(c) + Cos(a) * Sin(b) * Sin(c)
            Me.c = Cos(a) * Sin(b) * Cos(c) - Sin(a) * Cos(b) * Sin(c)
            Me.d = Cos(a) * Cos(b) * Sin(c) + Sin(a) * Sin(b) * Cos(c)
        Case Else
            Err.Raise 5, , "Convention de rotation non reconnue : " & convention
            Exit Sub
    End Select

End Sub


' Convertit un quaternion en angles d'Euler
Function toEulerAngle(isRad As Boolean) As Variant
    Dim angles(0 To 2) As Double
    Dim q0 As Double, q1 As Double, q2 As Double, q3 As Double
    Dim normilizedQ As Quaternion
    Set normilizedQ = Me.normalized
    
    q0 = normilizedQ.a
    q1 = normilizedQ.b
    q2 = normilizedQ.c
    q3 = normilizedQ.d

    angles(0) = WorksheetFunction.Atan2(2 * (q0 * q1 + q2 * q3), 1 - 2 * (q1 ^ 2 + q2 ^ 2))
    angles(1) = WorksheetFunction.Asin(2 * (q0 * q2 - q3 * q1))
    angles(2) = WorksheetFunction.Atan2(2 * (q0 * q3 + q1 * q2), 1 - 2 * (q2 ^ 2 + q3 ^ 2))

    ' Convertir les angles en degrés si nécessaire
    If Not isRad Then
        angles(0) = WorksheetFunction.Degrees(angles(0))
        angles(1) = WorksheetFunction.Degrees(angles(1))
        angles(2) = WorksheetFunction.Degrees(angles(2))
    End If

    toEulerAngle = angles
End Function


' MÃ©thode pour calculer la norme du Quaternion
Public Function getNorme() As Double
    getNorme = Sqr((Me.a) ^ 2 + (Me.b) ^ 2 + (Me.c) ^ 2 + (Me.d) ^ 2)
End Function

'Methode to normilize the Quaternion
Public Sub normalize()
    Dim norm As Double
    norm = getNorme()
    Me.QuaternionScale (1 / norm)
End Sub

' Retourne un quaternion Normalisé
Function normalized() As Quaternion
    Dim norm As Double
    Dim normalizedQ As New Quaternion
    normalizedQ.a = Me.a
    normalizedQ.b = Me.b
    normalizedQ.c = Me.c
    normalizedQ.d = Me.d
    
    ' Normilize
    normalizedQ.normalize
    
    'Normilize the Quaternion
    Set normalized = normalizedQ
End Function


' MÃ©thode pour additionner le Quaternion avec un autre Quaternion
Public Function Addition(q2 As Quaternion) As Quaternion
    Dim q3 As New Quaternion
    q3.a = Me.a + q2.a
    q3.b = Me.b + q2.b
    q3.c = Me.c + q2.c
    q3.d = Me.d + q2.d
    Set Addition = q3
End Function

' MÃ©thode pour multiplier le Quaternion par un scalaire
Public Function QuaternionScale(k As Double)
    Me.a = Me.a * k
    Me.b = Me.b * k
    Me.c = Me.c * k
    Me.d = Me.d * k
End Function

' MÃ©thode pour multiplier le Quaternion avec un autre Quaternion
Public Function quaternionMultiplication(q2 As Quaternion) As Quaternion
    Dim q3 As New Quaternion
    q3.a = Me.a * q2.a - Me.b * q2.b - Me.c * q2.c - Me.d * q2.d ' Composante rÃ©elle
    q3.b = Me.a * q2.b + Me.b * q2.a + Me.c * q2.d - Me.d * q2.c ' Composante imaginaire i
    q3.c = Me.a * q2.c - Me.b * q2.d + Me.c * q2.a + Me.d * q2.b ' Composante imaginaire j
    q3.d = Me.a * q2.d + Me.b * q2.c - Me.c * q2.b + Me.d * q2.a ' Composante imaginaire k
    Set quaternionMultiplication = q3
End Function

' Dans le module de classe Quaternion
Public Function getConjugate() As Quaternion
    Dim conj As New Quaternion
    conj.a = Me.a
    conj.b = -Me.b
    conj.c = -Me.c
    conj.d = -Me.d
    Set getConjugate = conj
End Function

Function transformPoint(q As Quaternion) As Quaternion
    ' Calculer le conjugué de q
    Dim q_conj As Quaternion
    Set q_conj = q.getConjugate()

    ' Calculer le nouveau point : new_p = q . Me . q-1
    Dim temp As Quaternion
    Set temp = q.quaternionMultiplication(Me)
    Set transformPoint = temp.quaternionMultiplication(q_conj)
End Function




