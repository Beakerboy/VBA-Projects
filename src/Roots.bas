Public Function CubeRoot(num)
    If num < 0 Then
        CubeRoot = -1 * (-1 * num) ^ (1 / 3)
    Else
        CubeRoot = num ^ (1 / 3)
    End If
End Function

Public Function Discriminant(a, b, c)
    Discriminant = b ^ 2 - (4 * a * c)
End Function

' Function: Quadratic
' Find the roots of a 2nd order polynomial
'
' Parameters:
'    a - x² coefficiant
'    b - x  coefficient
'    c - constant term
'
' Returns:
' An array of complex numbers of the form (r, i)
Public Function Quadratic(a, b, c) As Variant()

    Dim Results(1 To 2) As Variant
    Dim x1(1 To 2) As Variant
    Dim x2(1 To 2) As Variant
    If a = 0 Then
        Quadratic = -1 * c / b
    Else
        Descrim = Discriminant(a, b, c)
        If Descrim < 0 Then
            x1(1) = -1 * b / 2 / a
            x1(2) = -1 * Sqr(-1 * Descrim) / 2 / a
            x2(1) = -1 * b / 2 / a
            x2(2) = Sqr(-1 * Descrim) / 2 / a
        Else
            x1(1) = (-1 * b - Sqr(Descrim)) / 2 / a
            x1(2) = 0
            x2(1) = (-1 * b + Sqr(Descrim)) / 2 / a
            x2(2) = 0
        End If
    End If
    Results(1) = x1
    Results(2) = x2
    Quadratic = Results
End Function

' Function: Cubic
' Find the roots of a 3rd order polynomial
'
' Parameters:
'    a - x³ coefficiant
'    b - x² coefficient
'    c - x  coefficient
'    d - constant term
'
' Returns:
' An array of complex numbers of the form (r, i)
Public Function Cubic(a, b, c, d)
    Dim Results(1 To 3) As Variant
    Dim x1(1 To 2) As Variant
    Dim x2(1 To 2) As Variant
    Dim x3(1 To 2) As Variant
    If a = 0 Then
        Cubic = Quadratic(b, c, d)
    Else
        b = b / a
        c = c / a
        d = d / a
        a = 1
        ' Intermediate variables
        Q = (3 * c - b ^ 2) / 9
        R = (9 * b * c - 27 * d - 2 * b ^ 3) / 54
        
        ' Polynomial discriminant
        Descrim = Q ^ 3 + R ^ 2
        If Descrim < 0 Then
            'All unique real Roots
            Theta = Application.Acos(R / Sqr((-1 * Q) ^ 3))
            sqrt2q = 2 * Sqr(-1 * Q)
            p = Application.Pi()

            x1(1) = sqrt2q * Cos(Theta / 3) - (b / 3)
            x2(1) = sqrt2q * Cos((Theta + 2 * p) / 3) - (b / 3)
            x3(1) = sqrt2q * Cos((Theta + 4 * p) / 3) - (b / 3)
        Else
            S = CubeRoot(R + Sqr(Descrim))
            T = CubeRoot(R - Sqr(Descrim))
            If Abs(Descrim) < 0.000000000001 Then
                ' Real but with non-unique roots
                x1(1) = -b / 3 - (S + T) / 2
                x1(2) = 0
                x2(1) = S + T - b / 3
                x2(2) = 0
                x3(1) = -b / 3 - (S + T) / 2
                x3(2) = 0
            Else
                ' One Real, Two Complex
                x1(1) = S + T - b / 3
                x1(2) = 0
                quadb = b + x1(1)
                quadc = c + quadb + x1(1)
                Quad = Quadratic(1, quadb, quadc)
                x2(1) = Quad(1)(1)
                x2(2) = Quad(1)(2)
                x3(1) = Quad(2)(1)
                x3(2) = Quad(2)(2)
            End If
        End If
    End If
    Results(1) = x1
    Results(2) = x2
    Results(3) = x3
    Cubic = Results
End Function
