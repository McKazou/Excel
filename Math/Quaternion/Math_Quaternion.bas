Attribute VB_Name = "Math_Quaternion"
'Méthode pour obtenir le conjugué d'un quaternion
Public Function quaternionRotatePoint(p_range As Range, q_range As Range, Optional asArray As Boolean = True)
Attribute quaternionRotatePoint.VB_Description = "Méthode pour obtenir le conjugué d'un quaternion"
Attribute quaternionRotatePoint.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim point As New Quaternion
    Dim q As New Quaternion
    point.fromRange p_range
    q.fromRange q_range
    
    Dim point_result As New Quaternion
    Set point_result = point.transformPoint(q)
    
    If asArray Then
        quaternionRotatePoint = point_result.toArray
    Else
        quaternionRotatePoint = point_result.toString
    End If
End Function

' Méthode pour convertir un quaternion en un String
Public Function quaternionRangeToString(rng As Range, Optional numDecimals As Integer = -1)
Attribute quaternionRangeToString.VB_Description = "Méthode pour convertir un quaternion en un String"
Attribute quaternionRangeToString.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim q As New Quaternion
    q.fromRange rng
    quaternionRangeToString = q.toString(numDecimals)
End Function

' Méthode pour convertir une plage en chaîne quaternion
Public Function textToQuaternionArray(rng As Range)
Attribute textToQuaternionArray.VB_Description = "Méthode pour convertir une plage en chaîne quaternion"
Attribute textToQuaternionArray.VB_ProcData.VB_Invoke_Func = " \n20"
    'Not Implemented yet
    Dim q As New Quaternion
    q.fromString rng.Value2
    textToQuaternionArray = q.toArray
End Function

'Méthod pour obtenir le conjugé d'un quatenrion
Public Function quaternionGetConjugate(rng As Range, Optional asArray As Boolean = True) As Variant
Attribute quaternionGetConjugate.VB_Description = "Méthode pour obtenir le conjugé d'un quaternion"
Attribute quaternionGetConjugate.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim q As New Quaternion
    q.fromRange rng
    Set q = q.getConjugate
    
    If asArray Then
        quaternionGetConjugate = q.toArray
    Else
        quaternionGetConjugate = q.toString
    End If
End Function


'Methode pour multiplier deux Quaternions
Public Function quaternionMultiplication(rng1 As Range, rng2 As Range, Optional asArray As Boolean = True) As Variant
Attribute quaternionMultiplication.VB_Description = "Méthode pour multiplier deux Quaternions"
Attribute quaternionMultiplication.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim q1 As New Quaternion
    Dim q2 As New Quaternion
    Dim q3 As New Quaternion
    q1.fromRange rng1
    q2.fromRange rng2
    Set q3 = q1.quaternionMultiplication(q2)
    
    If asArray Then
        quaternionMultiplication = q3.toArray
    Else
        quaternionMultiplication = q3.toString
    End If
End Function

'Methode pour convertir un quaternio vers les angles d'Euler
Public Function quaternionToEulerAngle(rng As Range) As Variant
Attribute quaternionToEulerAngle.VB_Description = "Méthode pour convertir un quaternion vers les angles d'Euler"
Attribute quaternionToEulerAngle.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim q1 As New Quaternion
    q1.fromRange rng
    quaternionToEulerAngle = q1.toEulerAngle(False)
End Function

'Methode pour convertir les angles d'Euler vers un quaternion
Public Function eulerAngleToQuaternion(rng As Range, Optional convention As String = "ZYX") As Variant
Attribute eulerAngleToQuaternion.VB_Description = "Méthode pour convertir les angles d'Euler vers un quaternion"
Attribute eulerAngleToQuaternion.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim rngValues As Variant
    rngValues = rng.Value2
    
    ' Définir les angles d'Euler
    Dim phi As Double, theta As Double, psi As Double
    phi = CDbl(rngValues(1, 1))
    theta = CDbl(rngValues(2, 1))
    psi = CDbl(rngValues(3, 1))
    
    Dim q1 As New Quaternion
    q1.fromEulerAngle phi, theta, psi, False, convention
    eulerAngleToQuaternion = q1.toArray
End Function


'----------------------REGISTER MACRO FUNCTION -----------------
Sub RegisterQuaternionFunctions()
    Application.MacroOptions _
        Macro:="quaternionRotatePoint", _
        Description:="Méthode pour obtenir le conjugué d'un quaternion", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le point p", "Plage pour le quaternion q", "Option pour retourner le résultat en tant que tableau")
        
    Application.MacroOptions _
        Macro:="quaternionRangeToString", _
        Description:="Méthode pour convertir un quaternion en un String", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le quaternion", "Nombre optionnel de décimales")
        
    Application.MacroOptions _
        Macro:="textToQuaternionArray", _
        Description:="Méthode pour convertir une plage en chaîne quaternion", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le texte du quaternion")
        
    Application.MacroOptions _
        Macro:="quaternionGetConjugate", _
        Description:="Méthode pour obtenir le conjugé d'un quaternion", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le quaternion", "Option pour retourner le résultat en tant que tableau")
        
    Application.MacroOptions _
        Macro:="quaternionMultiplication", _
        Description:="Méthode pour multiplier deux Quaternions", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le premier quaternion", "Plage pour le deuxième quaternion", "Option pour retourner le résultat en tant que tableau")
        
    Application.MacroOptions _
        Macro:="quaternionToEulerAngle", _
        Description:="Méthode pour convertir un quaternion vers les angles d'Euler", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour le quaternion")
        
    Application.MacroOptions _
        Macro:="eulerAngleToQuaternion", _
        Description:="Méthode pour convertir les angles d'Euler vers un quaternion", _
        Category:="Quaternion Functions", _
        ArgumentDescriptions:=Array("Plage pour les angles d'Euler", "Convention optionnelle pour les angles d'Euler")
End Sub



' Fonction de test pour la classe Quaternion
Private Sub testQuaternion()
    Dim q As New Quaternion
    Dim s As String
    Dim arr() As Variant
    Dim testArr() As Variant
    Dim i As Integer

    ' Test de fromString et toString
    s = "-1-2i-3j-4k"
    q.fromString s
    Debug.Print "Test fromString/toString: " & (q.toString = s) & " output= " & q.toString ' Doit afficher "True"
    
    ' Test de fromString et toString
    s = "1+2i+3j+4k"
    q.fromString s
    Debug.Print "Test fromString/toString: " & (q.toString = s) & " output= " & q.toString  ' Doit afficher "True"
    
    ' Test de fromString et toString
    s = "-2i-4k"
    q.fromString s
    Debug.Print "Test fromString/toString: " & (q.toString = s) & " output= " & q.toString  ' Doit afficher "True"

    ' Test de fromArray et toArray
    ReDim arr(3)
    For i = 0 To 3
        arr(i) = i + 1
    Next i
    q.fromArray arr
    testArr = q.toArray
    For i = 0 To 3
        If arr(i) <> testArr(i) Then
            Debug.Print "Test fromArray/toArray: False"
            Exit Sub
        End If
    Next i
    Debug.Print "Test fromArray/toArray: True" ' Doit afficher "True"
    
    ' Test Normilize
    s = "-1-2i-3j-4k"
    Dim result As String
    result = "-0,182574185835055-0,365148371670111i-0,547722557505166j-0,730296743340221k"
    q.fromString s
    q.normalize
    Debug.Print "Test Normilize: " & (q.toString = result) & " output= " & q.toString  ' Doit afficher "True"
    
    
End Sub

