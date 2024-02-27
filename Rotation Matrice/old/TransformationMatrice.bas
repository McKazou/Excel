Attribute VB_Name = "TransformationMatrice"
Option Explicit
'Option Compare Binary


'This function will create the result matrice to transforme in point
'InputPoint:
' X
' Y
' Z
'Input Transformation:
'   Tx
'   Ty
'   Tz
'   AngleDegree rot X
'   AngleDegree rot Y
'   AngleDegree rot Z
'-----------------------------------
'Output :
'   R11     R12     R13     Tx
'   R21     R22     R23     Ty
'   R31     R32     R33     Tz
'   0       0       0       1

Function TransformPoint(ByVal InputRange As Range, Optional ByVal TranslationRange As Range, Optional ByVal RotationRange As Range, Optional ByVal inverse As Boolean) As Variant()
    
    Dim InputArray As Variant
    InputArray = getInputArray(InputRange)
    Dim transArray As Variant
    transArray = getTransArray(TranslationRange, RotationRange)
    
    Dim finalTransform As Variant
    finalTransform = getTransformationMatrix(transArray)
    
    If inverse Then
        finalTransform = WorksheetFunction.MInverse(finalTransform)
    End If
    
    PrintMatrix (finalTransform)
    Dim resultPoint As Variant
    'Application.WorksheetFunction.MMult
    'Application.Transpose
    resultPoint = Application.WorksheetFunction.MMult(finalTransform, Application.Transpose(InputArray))
    
    'For Each Val In resultPoint
    '    If Val < 0.000000001 Then
    '        Val = 0
    '    End If
    'Next
    
    TransformPoint = resultPoint

End Function

Function getTransformation(Optional ByVal TranslationRange As Range, Optional ByVal RotationRange As Range) As Variant
    
    Dim transArray As Variant
    transArray = getTransArray(TranslationRange, RotationRange)
    
    Dim finalTransform As Variant
    finalTransform = getTransformationMatrix(transArray)
    PrintMatrix (finalTransform)
    getTransformation = finalTransform
End Function

Private Function getInputArray(ByVal InputRange As Range)
    'Input as an array
    Dim InputArray As Variant 'Coordonnee given as an array
    ReDim InputArray(4 - 1)
    InputArray(3) = 1
    
    Dim cell As Range
    Dim cptInput As Integer

    '------------- HANDLE INPUT POINT -----------------------
    cptInput = 0
    'I need to look if the Inputrange is horizontal or vertical
    'MsgBox InputRange.Columns.Count & ":" & InputRange.Rows.Count
    If InputRange.Columns.Count = 3 And InputRange.Rows.Count = 1 _
    Or InputRange.Columns.Count = 4 And InputRange.Rows.Count = 1 Then
     'it's in on row

        For Each cell In InputRange
            InputArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
    ElseIf InputRange.Columns.Count = 1 And InputRange.Rows.Count = 3 _
    Or InputRange.Columns.Count = 1 And InputRange.Rows.Count = 4 Then
     'it's in one row
        For Each cell In InputRange
            InputArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
    Else
        'Error
        ReDim TransformationMatrix(1)
            Debug.Print "#InputPointNotGoodFormat"
            TransformationMatrix = CVErr(xlErrNA)
        Exit Function
    End If
    getInputArray = InputArray
End Function


Private Function getTransArray(ByVal TranslationRange As Range, ByVal RotationRange As Range)
    'Transformation as an array
    Dim transArray As Variant 'Coordonnee given as an array
    ReDim transArray(6 - 1)
    transArray(0) = 0
    transArray(1) = 0
    transArray(2) = 0
    transArray(3) = 0
    transArray(4) = 0
    transArray(5) = 0
    
    Dim cell As Range
    Dim cptInput As Integer
    
    '------------- HANDLE INPUT TRANSFORM -----------------------
    cptInput = 0
    'I need to look if the Inputrange is horizontal or vertical
    'MsgBox InputRange.Columns.Count & ":" & InputRange.Rows.Count
    If TranslationRange Is Nothing And RotationRange Is Nothing Then
    'Both parameters optional
            getTransArray = transArray
    ElseIf TranslationRange Is Nothing And Not RotationRange Is Nothing Then
        If (RotationRange.Columns.Count = 3 And RotationRange.Rows.Count = 1 _
        Or RotationRange.Columns.Count = 1 And RotationRange.Rows.Count = 3) Then
        'Only translation optional
            cptInput = 3
            For Each cell In RotationRange
                transArray(cptInput) = cell.Value
                cptInput = cptInput + 1
            Next
        End If
    ElseIf RotationRange Is Nothing And Not TranslationRange Is Nothing Then
        If (TranslationRange.Columns.Count = 3 And TranslationRange.Rows.Count = 1 _
        Or TranslationRange.Columns.Count = 1 And TranslationRange.Rows.Count = 3) Then
        'only rotation optional
            cptInput = 0
            For Each cell In TranslationRange
                transArray(cptInput) = cell.Value
                cptInput = cptInput + 1
            Next
        End If
    ElseIf TranslationRange.Columns.Count = 3 And TranslationRange.Rows.Count = 1 _
    And RotationRange.Columns.Count = 3 And RotationRange.Rows.Count = 1 Then
     'it's in on row
        For Each cell In TranslationRange
            transArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
        For Each cell In RotationRange
            transArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
    ElseIf TranslationRange.Columns.Count = 1 And TranslationRange.Rows.Count = 3 _
    And RotationRange.Columns.Count = 1 And RotationRange.Rows.Count = 3 Then
     'it's in one row
        For Each cell In TranslationRange
            transArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
        For Each cell In RotationRange
            transArray(cptInput) = cell.Value
            cptInput = cptInput + 1
        Next
    Else
        'Error
        ReDim TransformationMatrix(1)
           Debug.Print "#TransformPointNotGoodFormat"
           TransformationMatrix = CVErr(xlErrNA)
       Exit Function
    End If

    
    getTransArray = transArray
End Function



Private Function getTransformationMatrix(transArray As Variant) As Variant
        'Get Matrix Equivalent transformation
    Dim Tx, Ty, Tz As Double
    Dim rotX, rotY, rotZ As Double

    Tx = transArray(0)
    Ty = transArray(1)
    Tz = transArray(2)
    rotX = transArray(3)
    rotY = transArray(4)
    rotZ = transArray(5)

    Debug.Print ("Transformation extracted:")
    Debug.Print ("Translation - Tx:" & Tx & " | Ty: " & Ty & " | Tz: " & Tz)
    Debug.Print ("Rotation - rotX:" & rotX & " | rotY: " & rotY & " | rotZ: " & rotZ)

    Dim translate As Variant
    Dim rotateX, rotateY, rotateZ As Variant
    translate = Create_TranslationMatrix(Tx, Ty, Tz)
    rotateX = Create_RotationMatrix(rotX, "x")
    rotateY = Create_RotationMatrix(rotY, "y")
    rotateZ = Create_RotationMatrix(rotZ, "z")
    
    'Transformation final
    'Order : rotation X then rotation Y then rotation Z
    'Application.WorksheetFunction.MMult(xx,xx)
    Dim finalTransform As Variant
    finalTransform = Application.WorksheetFunction.MMult(rotateY, rotateX)
    finalTransform = Application.WorksheetFunction.MMult(rotateZ, finalTransform)
    finalTransform = Application.WorksheetFunction.MMult(translate, finalTransform)

    getTransformationMatrix = finalTransform
End Function


Private Function Create_RotationMatrix(ByVal Angle As Double, ByVal Axe As String) As Variant

'Create Rotation matrix

Angle = WorksheetFunction.Radians(Angle)

Dim Rotation As Variant
Rotation = Create_Matrix(4, 4)


Select Case Axe
    Case "X", "x"
    'Output :
    '   1       0       0      0
    '   0       CosA    -SinA  0
    '   0       SinA    CosA   0
    '   0       0       0      1
        'Diagonal
        Rotation(1, 1) = 1
        Rotation(4, 4) = 1

        'Rotation
        Rotation(2, 2) = Cos(Angle)
        Rotation(3, 3) = Cos(Angle)
        Rotation(2, 3) = -1 * Sin(Angle)
        Rotation(3, 2) = Sin(Angle)
    Case "Y", "y"
    'Output :
    '   CosA    0       sinA   0
    '   0       1       0      0
    '   -sinA   0       CosA   0
    '   0       0       0      1
        'Diagonal
        Rotation(2, 2) = 1
        Rotation(4, 4) = 1

        'Rotation
        Rotation(1, 1) = Cos(Angle)
        Rotation(3, 3) = Cos(Angle)
        Rotation(3, 1) = -1 * Sin(Angle)
        Rotation(1, 3) = Sin(Angle)
    Case "Z", "z"
    'Output :
    '   cosA    SinA    0      0
    '   -SinA   CosA    0      0
    '   0       0       1      0
    '   0       0       0      1
        'Diagonal
        Rotation(3, 3) = 1
        Rotation(4, 4) = 1

        'Rotation
        Rotation(1, 1) = Cos(Angle)
        Rotation(2, 2) = Cos(Angle)
        Rotation(2, 1) = Sin(Angle)
        Rotation(1, 2) = -1 * Sin(Angle)
    Case Default
        
End Select


Create_RotationMatrix = Rotation
Rotation = Null
End Function

Private Function Create_TranslationMatrix(ByVal Tx As Double, ByVal Ty As Double, ByVal Tz As Double) As Variant
'Create Translation matrix
'Output :
'   1       0       0     Tx
'   0       1       0     Ty
'   0       0       1     Tz
'   0       0       0      1
Dim Translation As Variant
Translation = Create_Matrix(4, 4)
'Diagonal
Translation(1, 1) = 1
Translation(2, 2) = 1
Translation(3, 3) = 1
Translation(4, 4) = 1

'Translation values
Translation(1, 4) = Tx
Translation(2, 4) = Ty
Translation(3, 4) = Tz

Create_TranslationMatrix = Translation
Translation = Null

End Function


Private Function Create_Matrix(x As Long, y As Long) As Variant
    
    Dim i, j As Integer
    Dim Arr() As Variant ' Matrix array
    ReDim Arr(1 To y, 1 To x)
    
    'Intialised at 0
    Dim Val As Variant
    Dim cpt As Integer
    cpt = 0
    For i = 1 To UBound(Arr, 1)

        For j = 1 To UBound(Arr, 2)

            Arr(i, j) = 0
            cpt = cpt + 1
        Next ' next column

    Next
    Create_Matrix = Arr

End Function


Private Sub PrintMatrix(Matrix As Variant)

    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To UBound(Matrix, 1)

        For j = 1 To UBound(Matrix, 2)

            Debug.Print Matrix(i, j);
        Next ' next column
        Debug.Print ""

    Next
End Sub




