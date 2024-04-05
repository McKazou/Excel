Attribute VB_Name = "fftProgram"
'Module: fftProgram
'Author: HJ Park
'Date  : 2019.5.18(v1.0), 2022.8.1(v2.0)
'https://infograph.tistory.com/351

Option Explicit

Public Const myPI As Double = 3.14159265358979

Public Function Log2(X As Long) As Double
  Log2 = Log(X) / Log(2)
End Function

Public Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function

Public Function Floor(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' Factor is the multiple to which you want to round
    Floor = Int(X / Factor) * Factor
End Function

' return 0 if N is 2^n value,
' return (2^n - N) if N is not 2^n value. 2^n is Ceiling value.
' return -1, if error
Public Function IsPowerOfTwo(N As Long) As Long
  If N = 0 Then GoTo EXIT_FUNCTION
  
  Dim c As Long, F As Double
  c = Ceiling(Log2(N)) 'Factor=0, therefore C is an integer number
  F = Floor(Log2(N))
  
  If c = F Then
    IsPowerOfTwo = 0
  Else
    IsPowerOfTwo = (2 ^ c - N)
  End If
  Exit Function
  
EXIT_FUNCTION:
  IsPowerOfTwo = -1
End Function




''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''

Function MakePowerOfTwoSize(ByRef r As Range, ByVal fillCount As Long) As Boolean
  Dim arr() As Integer
  On Error GoTo ERROR_HANDLE
  
  '1)make a array with zero
  ReDim arr(0 To fillCount - 1) As Integer
    
  '2)set a range to be filled with zero
  Dim fillRowStart As Long
  Dim fillRange As Range
  
  fillRowStart = r.Row + r.Rows.count
  Set fillRange = Range(Cells(fillRowStart, r.Column), Cells(fillRowStart + fillCount - 1, r.Column))
  
  '3)fill as zero
  fillRange = arr
  
  '4)update range area to be extended
  Set r = Union(r, fillRange)
  
  MakePowerOfTwoSize = True
  Exit Function
  
ERROR_HANDLE:
  MakePowerOfTwoSize = False
End Function

' read the range and return it as complex value array
Function Range2Array(r As Range) As Complex()
  Dim i As Long, size As Long
  Dim arr() As Complex
  
  size = r.Rows.count
  ReDim arr(0 To size - 1) As Complex
  
  Dim re As Double, im As Double
  
  On Error GoTo ERROR_HANDLE
  For i = 1 To size
    arr(i - 1) = String2Complex(r.Rows(i).Value)
  Next i
  
  Range2Array = arr
  
  Exit Function
ERROR_HANDLE:
  MsgBox "Error: " & i
End Function

Function ArrangedNum(num As Long, numOfBits As Integer) As Long
  Dim arr() As Byte
  Dim i As Integer, j As Integer
  Dim k As Long
  
  If (2 ^ numOfBits) <= num Then GoTo EXIT_FUNCTION
  
  '1) Decimal number -> Reversed Binary array : (13,4) -> {1,1,0,1} -> {1,0,1,1}
  ReDim arr(0 To numOfBits - 1) As Byte
  For i = 0 To numOfBits - 1
    j = (numOfBits - 1) - i
    k = Int((num / (2 ^ j)))
    arr(j) = (k And 1)
  Next i
  
  '2) Reversed Binary -> Decimal: {1,0,1,1} -> 1*2^3 + 0*2^2 + 1*2&1 + 1 = 11
  Dim d As Long
  For i = 0 To numOfBits - 1
    d = d + (arr(i) * 2 ^ (numOfBits - 1 - i))
  Next i
  
  ArrangedNum = d
  Exit Function
  
EXIT_FUNCTION:
  ArrangedNum = 0
End Function

' rangeArr[1 to n, 1]
Function arrangeToFFTArray(arr() As Complex, size As Long, numOfBits As Integer) As Complex()
  Dim i As Long, j As Long
  Dim arrangedArr() As Complex
    
  ReDim arrangedArr(0 To size - 1) As Complex
  For i = 0 To size - 1
    j = ArrangedNum(i, numOfBits)  '{000,001,010, 011, 100, 101, 110, 111} -> {0, 4, 2, 6, 1, 5, 3, 7}
    arrangedArr(j) = arr(i)
  Next i
  
  arrangeToFFTArray = arrangedArr
End Function

' calculate convolution ring W
' W[k] = cos(theta) - isin(theta)
'   theta = (2pi*k/N)
Function CalculateW(cnt As Long, isInverse As Boolean) As Complex()
  Dim arr() As Complex
  Dim i As Long
  Dim T As Double, theta As Double
  Dim N As Long, N2 As Long
  
  N = cnt
  N2 = N / 2
  ReDim arr(0 To N2 - 1) As Complex  'enough to calculate 0 to (N/2 -1)
  T = 2 * myPI / CDbl(N)
  
  If isInverse Then
    For i = 0 To N2 - 1
      theta = -(T * i)
      arr(i) = Cplx(Cos(theta), -Sin(theta))
    Next i
  Else
    For i = 0 To N2 - 1
      theta = T * i
      arr(i) = Cplx(Cos(theta), -Sin(theta))
    Next i
  End If
  
  CalculateW = arr
End Function

' X({0,1}, [0,n-1]): 2d array.  (0, n) <--> (1,n)
' src: src index of the array. 0 or 1
' tgt: tgt index of the array. 1 or 0
' s : starting index of the data in the array
' size: region size to be calculated
' kJump : k's jumping value
' W(0 ~ n-1) : Convolution ring
Sub RegionFFT(X() As Complex, src As Integer, tgt As Integer, _
            s As Long, size As Long, kJump As Long, W() As Complex)
  Dim i As Long, e As Long
  Dim half As Long
  Dim k As Long
  Dim T As Complex
  
  ' Xm+1[i] = Xm[i] + Xm[i+half]W[k]
  ' Xm+1[i+half] = Xm[i] - Xm[i+half]W[k]
  k = 0
  e = s + (size / 2) - 1
  half = size / 2
  For i = s To e
    T = CMult(X(src, i + half), W(k))
    X(tgt, i) = CAdd(X(src, i), T)
    X(tgt, i + half) = CSub(X(src, i), T)
    k = k + kJump
  Next i
End Sub

Sub WriteToTarget(tgtRange As Range, X() As Complex, tgtIdx As Integer, N As Long, roundDigit As Integer)
  Dim i As Long
  Dim arr() As Variant
  
  ReDim arr(0 To N - 1) As Variant
  For i = 0 To N - 1
   If X(tgtIdx, i).im < 0 Then
     arr(i) = Round(X(tgtIdx, i).re, roundDigit) & Round(X(tgtIdx, i).im, roundDigit) & "i"
   Else
     arr(i) = Round(X(tgtIdx, i).re, roundDigit) & "+" & Round(X(tgtIdx, i).im, roundDigit) & "i"
   End If
  Next i
  
  tgtRange.Rows = Application.Transpose(arr)
End Sub

' xRange: input data
' tgtRange: output range
' isInverse: FFT or IFFT
Public Function FFT_Forward(xRange As Range, tgtRangeStart As Range, roundDigit As Integer, isInverse As Boolean) As Complex()
  Dim i As Long, N As Long
  Dim totalLoop As Integer, curLoop As Integer 'enough as Integer b/c it is used for loop varoable
  Dim xArr() As Complex, xSortedArr() As Complex
  Dim W() As Complex 'convolution ring
  Dim X() As Complex 'output result
  Dim errMsg As String
  
  errMsg = "Uncatched error"
    
  '1) check whether 2^r count data, if not pad to zero
  Dim fillCount As Long
  N = xRange.Rows.count
  fillCount = IsPowerOfTwo(N)
  If fillCount = -1 Then
    errMsg = "No input data. Choose input data"
    GoTo ERROR_HANDLE
  End If
  If fillCount <> 0 Then
    If MakePowerOfTwoSize(xRange, fillCount) = False Then  'xRange's size will be chnaged
      errMsg = "Error while zero padding"
      GoTo ERROR_HANDLE
    End If
  End If
  
  '2) calculate loop count for FFT: 2->1  4->2  8->3 ...
  N = xRange.Rows.count 'xRange's size can be changed so read one more...
  totalLoop = Log2(N)
  
  '3) sort x for 2's FFT : convert to reversed binary and then convert to decimal
  xArr = Range2Array(xRange)  'xArr[0,n-1]
  xSortedArr = arrangeToFFTArray(xArr, N, totalLoop) 'xSortedArr[0,n-1]
  
  '4) calculate W
  W = CalculateW(N, isInverse)
  
  '5) use 2-dimensional array to save memory space. X[0, ] <-> X[1, ]
  ReDim X(0 To 1, 0 To N - 1) As Complex
  For i = 0 To N - 1
    X(0, i) = xSortedArr(i)
  Next i
  
  '6) Do 2's FFT with sorted x
  Dim srcIdx As Integer, tgtIdx As Integer
  Dim kJump As Long, regionSize As Long
  
  tgtIdx = 0
  For curLoop = 0 To totalLoop - 1
    tgtIdx = (tgtIdx + 1) Mod 2
    srcIdx = (tgtIdx + 1) Mod 2
    regionSize = 2 ^ (curLoop + 1)          ' if N=8: 2 -> 4 -> 8
    kJump = 2 ^ (totalLoop - curLoop - 1)    ' if N=8: 4 -> 2 -> 1
    i = 0
    Do While i < N
      Call RegionFFT(X, srcIdx, tgtIdx, i, regionSize, kJump, W)
      i = i + regionSize
    Loop
  Next curLoop
  
  '7)return the value
  Dim resultIdx As Integer
  If (totalLoop Mod 2) = 0 Then resultIdx = 0 Else resultIdx = 1
  
  Dim result() As Complex
  ReDim result(0 To N - 1) As Complex
  If isInverse = True Then
    For i = 0 To N - 1
      result(i) = CDivR(X(resultIdx, i), N)
    Next i
  Else
    For i = 0 To N - 1
      result(i) = X(resultIdx, i)
    Next i
  End If
  
  FFT_Forward = result
  
  Exit Function
  
ERROR_HANDLE:
  Err.Raise Number:=vbObjectError, Description:=("FFT calculation error: " & errMsg)
End Function


Public Sub FFT(xRange As Range, tgtRangeStart As Range, roundDigit As Integer)
  Dim X() As Complex
  Dim tgtRange As Range
  
  '1. calculate FFT_forward value
  On Error GoTo ERROR_HANDLE
  X = FFT_Forward(xRange, tgtRangeStart, roundDigit, False)
  
  '2. write to the worksheet
  Dim N As Long
  N = UBound(X) - LBound(X) + 1
  
  Dim i As Long
  Dim arr() As Variant
  ReDim arr(0 To N - 1) As Variant
  For i = 0 To N - 1
   If X(i).im < 0 Then
     arr(i) = Round(X(i).re, roundDigit) & Round(X(i).im, roundDigit) & "i"
   Else
     arr(i) = Round(X(i).re, roundDigit) & "+" & Round(X(i).im, roundDigit) & "i"
   End If
  Next i
  
  Set tgtRange = Range(Cells(tgtRangeStart.Row, tgtRangeStart.Column), Cells(tgtRangeStart.Row + N - 1, tgtRangeStart.Column))
  tgtRange.Rows = Application.Transpose(arr)
  Exit Sub
  
ERROR_HANDLE:
  
End Sub

Public Sub IFFT(xRange As Range, tgtRangeStart As Range, roundDigit As Integer)
  Dim X() As Complex
  Dim tgtRange As Range
  
  '1. calculate FFT_forward value
  On Error GoTo ERROR_HANDLE
  X = FFT_Forward(xRange, tgtRangeStart, roundDigit, True)
  
  '2.write to the worksheet
  Dim N As Long
  N = UBound(X) - LBound(X) + 1
  
  Dim arr() As Variant
  ReDim arr(0 To N - 1) As Variant
  Dim i As Long
  For i = 0 To N - 1
    arr(i) = Round(X(i).re, roundDigit)
  Next i
  
  Set tgtRange = Range(Cells(tgtRangeStart.Row, tgtRangeStart.Column), Cells(tgtRangeStart.Row + N - 1, tgtRangeStart.Column))
  tgtRange.Rows = Application.Transpose(arr)
  Exit Sub
  
ERROR_HANDLE:

End Sub

'Be carefull with this one it can take a long time :
Public Function FourrierTransformation(xRange As Range, roundDigit As Integer) As Variant
  Dim X() As Complex
  Dim N As Long
  Dim i As Long
  Dim arr() As Variant

  '1. calculate FFT_forward value
  On Error GoTo ERROR_HANDLE
  X = FFT_Forward(xRange, Nothing, roundDigit, False)

  '2. write to the array
  N = UBound(X) - LBound(X) + 1
  ReDim arr(0 To N - 1) As Variant
  For i = 0 To N - 1
   If X(i).im < 0 Then
     arr(i) = Round(X(i).re, roundDigit) & Round(X(i).im, roundDigit) & "i"
   Else
     arr(i) = Round(X(i).re, roundDigit) & "+" & Round(X(i).im, roundDigit) & "i"
   End If
  Next i

  FourrierTransformation = arr
  Exit Function

ERROR_HANDLE:
  Err.Raise Number:=vbObjectError, Description:=("FFT calculation error: " & "Uncatched error")
End Function

Public Function FourrierTransformationInverse(xRange As Range, roundDigit As Integer) As Variant
  Dim X() As Complex
  Dim N As Long
  Dim i As Long
  Dim arr() As Variant

  '1. calculate FFT_forward value with isInverse set to True
  On Error GoTo ERROR_HANDLE
  X = FFT_Forward(xRange, Nothing, roundDigit, True)

  '2. write to the array
  N = UBound(X) - LBound(X) + 1
  ReDim arr(0 To N - 1) As Variant
  For i = 0 To N - 1
    arr(i) = Round(X(i).re, roundDigit)
  Next i

  FourrierTransformationInverse = arr
  Exit Function

ERROR_HANDLE:
  Err.Raise Number:=vbObjectError, Description:=("IFFT calculation error: " & "Uncatched error")
End Function


Sub LoadFFTForm()
  FFT_Form.Show
End Sub

