Attribute VB_Name = "RotationMatriceCalculation"
Option Explicit

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

'This doesn't work and i'm bored to fix it
Public Function ChainTransformationPoint(Optional ByVal inputRange As Range, Optional ByVal MultipleTranslationRange As Range, Optional ByVal MultipleRotationRange As Range, Optional ByVal inverse As Boolean) As Variant
    'This will try to solv the fact that to calculate a chanin of transformation i need to cumulate all previsous rotation

    'The InputRange should be a vector point (X,Y,Z)
        Dim inputArray As Variant
        inputArray = getInputArray(inputRange)

     'The MultipleTranslationRange should be a range of dimension (3,x) or (x,3) : THIS WILL COME LATER

     'The MultipleRotationRange should be a range of dimension (3,x) or (x,3) : For each x we will multiple the transformation matrix in order to get the final matrix rotation
        Dim RotationList As Dictionary
        Set RotationList = RotationListToDic(MultipleRotationRange)

        Dim a, rotInList
        For Each a In RotationList
            'TransformPoint(inputRange,
            
            rotInList = RotationList.Item(a)
            inputArray = TransformPoint(inputArray, , rotInList, inverse)
        Next
End Function



Public Function TransformPoint(Optional ByVal inputRange As Variant, Optional ByVal TranslationRange As Variant, Optional ByVal RotationRange As Variant, Optional ByVal inverse As Boolean) As Variant
    
    Dim inputArray As Variant
    If TypeOf inputRange Is Range Then
        inputArray = getInputArray(inputRange)
    ElseIf IsMissing(inputRange) Then
        inputArray = getInputArray()
    End If
    
    Dim TransAsArray, rotAsArray As Variant
    TransAsArray = TranslationRange
    rotAsArray = RotationRange
    If TypeOf TranslationRange Is Range Then
        TransAsArray = TranslationRange.Value2
    End If
    If TypeOf RotationRange Is Range Then
        rotAsArray = RotationRange.Value2
        'is range
    End If

    Dim finalTransform As Variant
    If IsMissing(TranslationRange) And IsMissing(RotationRange) Then
            finalTransform = getTransformation()
    ElseIf Not IsMissing(TranslationRange) And IsMissing(RotationRange) Then
            finalTransform = getTransformation(TransAsArray)
    ElseIf IsMissing(TranslationRange) And Not IsMissing(RotationRange) Then
            finalTransform = getTransformation(, rotAsArray)
    ElseIf Not IsMissing(TranslationRange) And Not IsMissing(RotationRange) Then
            finalTransform = getTransformation(TransAsArray, rotAsArray)
    End If
    
    If inverse Then
        finalTransform = WorksheetFunction.MInverse(finalTransform)
    End If
    
    'PrintMatrix (finalTransform)
    Dim resultPoint As Variant
    'Application.WorksheetFunction.MMult
    'Application.Transpose
    'Stop
    resultPoint = Application.WorksheetFunction.MMult(finalTransform, Application.transpose(inputArray))
    
    TransformPoint = resultPoint

End Function

Function getTransformation(Optional ByVal TranslationRange As Variant, Optional ByVal RotationRange As Variant) As Variant

    Dim transArray As Variant
    Dim TransAsArray, rotAsArray As Variant
    TransAsArray = TranslationRange
    rotAsArray = RotationRange
    If TypeOf TranslationRange Is Range Then
        TransAsArray = TranslationRange.Value2
    End If
    If TypeOf RotationRange Is Range Then
        rotAsArray = RotationRange.Value2
        'is range
    End If
    
    If IsMissing(TranslationRange) And IsMissing(RotationRange) Then
            transArray = getTransArray()
    ElseIf Not IsMissing(TranslationRange) And IsMissing(RotationRange) Then
            transArray = getTransArray(TransAsArray)
    ElseIf IsMissing(TranslationRange) And Not IsMissing(RotationRange) Then
            transArray = getTransArray(, rotAsArray)
    ElseIf Not IsMissing(TranslationRange) And Not IsMissing(RotationRange) Then
            transArray = getTransArray(TransAsArray, rotAsArray)
    End If

    
    Dim finalTransform As Variant
    finalTransform = getTransformationMatrix(transArray)
    'PrintMatrix (finalTransform)
    getTransformation = finalTransform
End Function

Private Function getInputArray(Optional ByVal inputRange As Range)
    'Input as an array
    Dim inputArray As Variant 'Coordonnee given as an array
    ReDim inputArray(4 - 1)
    inputArray(0) = 0
    inputArray(1) = 0
    inputArray(2) = 0
    inputArray(3) = 1
    
    Dim cell As Range
    Dim cptInput As Integer

    If Not inputRange Is Nothing Then
        '------------- HANDLE INPUT POINT -----------------------
        cptInput = 0
        'I need to look if the Inputrange is horizontal or vertical
        'MsgBox InputRange.Columns.Count & ":" & InputRange.Rows.Count
        If inputRange.Columns.Count = 3 And inputRange.Rows.Count = 1 _
        Or inputRange.Columns.Count = 4 And inputRange.Rows.Count = 1 Then
         'it's in on row
    
            For Each cell In inputRange
                inputArray(cptInput) = cell.Value
                cptInput = cptInput + 1
            Next
        ElseIf inputRange.Columns.Count = 1 And inputRange.Rows.Count = 3 _
        Or inputRange.Columns.Count = 1 And inputRange.Rows.Count = 4 Then
         'it's in one row
            For Each cell In inputRange
                inputArray(cptInput) = cell.Value
                cptInput = cptInput + 1
            Next
        Else
            'Error
            ReDim TransformationMatrix(1)
                Debug.Print "#######InputPointNotGoodFormat#######"
                TransformationMatrix = CVErr(xlErrNA)
            Exit Function
        End If
    End If
    getInputArray = inputArray
End Function

Private Function RotationListToDic(inputRange As Range) As Dictionary
    
    Dim rowsNb As Integer
    Dim colsNb As Integer
    Dim inputArray As Variant
    inputArray = inputRange.Value2
    rowsNb = UBound(inputArray)
    colsNb = UBound(inputArray, 2)
    
    Dim dic As Dictionary                  'Create a variable
    Set dic = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    Dim slice(1 To 1, 1 To 3) As Double
        
    If colsNb = 3 And rowsNb <> 3 Then
        'Nothing to do
        
        For i = 1 To rowsNb
            slice(1, 1) = inputArray(i, 1)
            slice(1, 2) = inputArray(i, 2)
            slice(1, 3) = inputArray(i, 3)
            dic.Add i, slice
        Next
        Set RotationListToDic = dic
    ElseIf rowsNb = 3 And colsNb <> 3 Then
        'Transpose
        For i = 0 To colsNb
            slice(1, 1) = inputArray(1, i)
            slice(1, 2) = inputArray(2, i)
            slice(1, 3) = inputArray(3, i)
            dic.Add i, slice
        Next
        
        Set RotationListToDic = dic
    ElseIf rowsNb = colsNb And rowsNb = 3 Then
        'Cannnot know what direction so using the default one
        For i = 0 To rowsNb
            slice(1, 1) = inputArray(i, 1)
            slice(1, 2) = inputArray(i, 2)
            slice(1, 3) = inputArray(i, 2)
            dic.Add i, slice
        Next
        Set RotationListToDic = dic
    Else
        Debug.Print "[RotationListToDic] - ERROR IN INPUT"
        'RotationListToDic = CVErr(xlErrNA)
    End If

End Function

Private Function getTransArray(Optional ByVal TranslationArray As Variant, Optional ByVal RotationArray As Variant)
    'Transformation as an array
    Dim transArray As Variant 'Coordonnee given as an array
    ReDim transArray(6 - 1)
    transArray(0) = 0
    transArray(1) = 0
    transArray(2) = 0
    transArray(3) = 0
    transArray(4) = 0
    transArray(5) = 0
    
    Dim cell As Variant
    Dim cptInput As Integer
    
    '------------- HANDLE INPUT TRANSFORM -----------------------
    Dim TranslationColNb, TranslationRowNb As Integer
    Dim RotationColNb, RotationRowNb As Integer
    
        cptInput = 0
        
    If Not IsMissing(TranslationArray) Then
        'If Not TranslationArray = Empty Then
    
            TranslationRowNb = UBound(TranslationArray)
            TranslationColNb = UBound(TranslationArray, 2)
            
            If (TranslationColNb = 3 And TranslationRowNb = 1 _
            Or TranslationColNb = 1 And TranslationRowNb = 3) Then
            'only rotation optional
                cptInput = 0
                For Each cell In TranslationArray
                    If IsNumeric(cell) Then
                        transArray(cptInput) = cell
                        cptInput = cptInput + 1
                    Else
                    'Error probably not a value in the input array
                    End If
                Next
            End If
        'End If
    End If
    If Not IsMissing(RotationArray) Then
        'If Not RotationArray = Empty Then
    
            RotationRowNb = UBound(RotationArray)
            RotationColNb = UBound(RotationArray, 2)
            
            If (RotationColNb = 3 And RotationRowNb = 1 _
            Or RotationColNb = 1 And RotationRowNb = 3) Then
            'Only translation optional
                cptInput = 3
                For Each cell In RotationArray
                    If IsNumeric(cell) Then
                        transArray(cptInput) = cell
                        cptInput = cptInput + 1
                    Else
                    'Error probably not a value in the input array
                    End If
                Next
            End If
        'End If
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
    
    Dim Val As Double

    'Debug.Print ("Transformation extracted:")
    'Debug.Print ("Translation - Tx:" & Tx & " | Ty: " & Ty & " | Tz: " & Tz)
    'Debug.Print ("Rotation - rotX:" & rotX & " | rotY: " & rotY & " | rotZ: " & rotZ)

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
            Rotation(2, 2) = Round(Cos(Angle), 8)
            Rotation(3, 3) = Round(Cos(Angle), 8)
            Rotation(2, 3) = Round(-1 * Sin(Angle), 8)
            Rotation(3, 2) = Round(Sin(Angle), 8)
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
            Rotation(1, 1) = Round(Cos(Angle), 8)
            Rotation(3, 3) = Round(Cos(Angle), 8)
            Rotation(3, 1) = Round(-1 * Sin(Angle), 8)
            Rotation(1, 3) = Round(Sin(Angle), 8)
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
            Rotation(1, 1) = Round(Cos(Angle), 8)
            Rotation(2, 2) = Round(Cos(Angle), 8)
            Rotation(2, 1) = Round(Sin(Angle), 8)
            Rotation(1, 2) = Round(-1 * Sin(Angle), 8)
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


Function TransposeArray(MyArray As Variant) As Variant
    Dim x As Long, y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    
    Dim tempArr As Variant
    
    'Get Upper and Lower Bounds
    maxX = UBound(MyArray, 1)
    minX = LBound(MyArray, 1)
    maxY = UBound(MyArray, 2)
    minY = LBound(MyArray, 2)
    
    'Create New Temp Array
    ReDim tempArr(minY To maxY, minX To maxX)
    
    'Transpose the Array
    For x = minX To maxX
        For y = minY To maxY
            tempArr(y, x) = MyArray(x, y)
        Next y
    Next x
    
    'Output Array
    TransposeArray = tempArr
    
End Function




