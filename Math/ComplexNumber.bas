'Module: complexNum

Option Explicit

Public Type Complex
    re As Double
    im As Double
End Type

Public Function Cplx(ByVal real As Double, ByVal imaginary As Double) As Complex
    Cplx.re = real
    Cplx.im = imaginary
End Function

Public Function CAdd(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    CAdd.re = z1.re + z2.re
    CAdd.im = z1.im + z2.im
End Function

Public Function CSub(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    CSub.re = z1.re - z2.re
    CSub.im = z1.im - z2.im
End Function

Public Function CMult(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    CMult.re = z1.re * z2.re - z1.im * z2.im
    CMult.im = z1.re * z2.im + z1.im * z2.re
End Function

Public Function CDivR(ByRef z As Complex, ByVal r As Double) As Complex
' Divide complex number by real number
    CDivR.re = z.re / r
    CDivR.im = z.im / r
End Function

Public Function CMagnitude(ByRef z As Complex) As Double
    CMagnitude = Sqr((z.re ^ 2) + (z.im ^ 2))
End Function

Public Function String2Complex(ByVal s As String) As Complex
  Dim sLen As Integer
  Dim re As Double, im As Double
  
  s = Trim(s)
  sLen = Len(s)
  
  If sLen = 0 Then: GoTo ERROR_HANDLE
  
  'support a+aj style
  If InStr(s, "j") Then
    s = Replace(s, "j", "i")
  End If
  
  'total cases: (1){a, -a}  (2){-a+bi  +a+bi} {-a-bi  +a-bi} {-bi  +bi} (3){a+bi, a-bi, bi}
  Dim pos As Integer
  If InStr(s, "i") = 0 Then ' (1){a, -a}
    re = CDbl(s)
    im = 0
  ElseIf InStr(s, "+") = 1 Or InStr(s, "-") = 1 Then '(2){-a+bi  +a+bi} {-a-bi  +a-bi} {-bi  +bi}
    pos = InStr(2, s, "+")
    If pos > 0 Then '-a+bi  +a+bi
      re = CDbl(Mid(s, 1, pos - 1))
      im = CDbl(Mid(s, pos + 1, sLen - pos - 1))
    Else
      pos = InStr(2, s, "-")
      If pos > 0 Then '-a-bi  +a-bi
        pos = InStr(2, s, "-")
        re = CDbl(Mid(s, 1, pos - 1))
        im = CDbl(Mid(s, pos, sLen - pos))
      Else ' -bi  +bi
        re = 0
        im = CDbl(Left(s, sLen - 1))
      End If
    End If
  Else '(3){a+bi a-bi, bi}
    pos = InStr(s, "+")
    If pos > 0 Then 'a+bi
      re = CDbl(Mid(s, 1, pos - 1))
      im = CDbl(Mid(s, pos + 1, sLen - pos - 1))
    Else 'a-bi  bi
      pos = InStr(s, "-")
      If pos > 0 Then 'a-bi
        re = CDbl(Mid(s, 1, pos - 1))
        im = CDbl(Mid(s, pos, sLen - pos))
      Else 'bi
        re = 0
        im = CDbl(Left(s, sLen - 1))
      End If
    End If
  End If
      
  String2Complex = Cplx(re, im)
  Exit Function
  
ERROR_HANDLE:
  String2Complex = Cplx(0, 0)
End Function