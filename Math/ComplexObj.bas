Class Complex
    Private pRe As Double
    Private pIm As Double
    Private pR As Double
    Private pTheta As Double

    ' Propriétés pour la partie réelle et imaginaire
    Public Property Get Re() As Double
        Re = pRe
    End Property

    Public Property Let Re(value As Double)
        pRe = value
        pR = Sqr((pRe ^ 2) + (pIm ^ 2))
        pTheta = Atn2(pIm, pRe)
    End Property

    Public Property Get Im() As Double
        Im = pIm
    End Property

    Public Property Let Im(value As Double)
        pIm = value
        pR = Sqr((pRe ^ 2) + (pIm ^ 2))
        pTheta = Atn2(pIm, pRe)
    End Property

    ' Propriétés pour le module et l'argument
    Public Property Get R() As Double
        R = pR
    End Property

    Public Property Let R(value As Double)
        pR = value
        pRe = pR * Cos(pTheta)
        pIm = pR * Sin(pTheta)
    End Property

    Public Property Get Theta() As Double
        Theta = pTheta
    End Property

    Public Property Let Theta(value As Double)
        pTheta = value
        pRe = pR * Cos(pTheta)
        pIm = pR * Sin(pTheta)
    End Property

    ' Autres méthodes pour les opérations complexes
    ' ...
    'CMagnitude could be here
    Public Function CMagnitude() As Double
        CMagnitude = Sqr((Me.re ^ 2) + (Me.im ^ 2))
    End Function
End Class


Public Function CAdd(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As New Complex
    result.Re = z1.Re + z2.Re
    result.Im = z1.Im + z2.Im
    Set CAdd = result
End Function

Public Function CSub(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As New Complex
    result.Re = z1.Re - z2.Re
    result.Im = z1.Im - z2.Im
    Set CSub = result
End Function

Public Function CMult(ByRef z1 As Complex, ByRef z2 As Complex) As Complex
    Dim result As New Complex
    result.Re = z1.re * z2.re - z1.im * z2.im
    result.Im = z1.re * z2.im + z1.im * z2.re
    Set CMult = result
End Function

Public Function CDivR(ByRef z As Complex, ByVal r As Double) As Complex
' Divide complex number by real number
    Dim result As New Complex
    result.Re = z.re / r
    result.Im = z.im / r
    Set CDivR = result
End Function

Public Function String2Complex(ByVal s As String) As Complex
  Dim sLen As Integer
  Dim re As Double, im As Double
  Dim r as double, theta as double
  Dim pos as Integer
  Dim complex as New Complex
  
  s = Trim(s)
  sLen = Len(s)
  
  If sLen = 0 Then
        Err.Raise Number:=vbObjectError + 9999, _
              Source:="Complex::String2Complex", _
              Description:="Cannot convert empty string to Complex."
  end if
  
  'support a+aj style
  If InStr(s, "j") Then
    s = Replace(s, "j", "i")
  End If

   ' Check if the string is in exponential form
  If InStr(s, "exp") > 0 Or InStr(s, "e^") > 0 Then
        s = Replace(s, "exp", "")
        s = Replace(s, "e^", "")
        s = Replace(s, "*", "")
        pos = InStr(s, "i")
        If pos > 0 Then
            r = CDbl(Left(s, pos - 1))
            theta = CDbl(Mid(s, pos + 1))
            complex.R = r
            complex.Theta = theta
        Else
            Err.Raise Number:=vbObjectError + 9999, _
              Source:="Complex::String2Complex", _
              Description:="Cannot convert string '" & s & "' to Complex."
        End If
    else
  
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
    End if
      complex.re = re
      complex.im = im 
    set String2Complex = complex
  Exit Function
  
ERROR_HANDLE:
    Dim complexeZero as new Complex
    complexeZero.re = 0
    complexeZero.im = 0
    set String2Complex = complexeZero
End Function