' Function: Diagonal
' Produces a diagonal matrix from a vector
'
' Parameters:
'   vector - An array of values
Private Function Diagonal(vector As Variant)
    Dim Arr() As Variant
    Dim R As Integer
    'If IsArray(vector) = True Then
    R = UBound(vector)
    'Else
    'R = vector.count
    'End If
    ReDim Arr(R, R)
    Dim i, j As Integer
    For i = 0 To R
        For j = 0 To R
            If j = i Then
                Arr(i, j) = vector(i)
            Else
                Arr(i, j) = 0
            End If
        Next j
    Next i
    Diagonal = Arr
End Function

' Function: Vander
' Create the Vandermonde Matrix whose next to
' the last column is 'vector' with a total width of 'width'
' The vandermonde matrix is:
'
'  vector(0)^width...vector(0)^2 vector(0) 1
'  vector(1)^width...vector(1)^2 vector(1) 1
'        .                 .           .   .
'        .                 .           .   .
'        .                 .           .   .
'  vector(n)^width...vector(n)^2 vector(n) 1
Public Function Vander(vector As Variant, width As Integer)
    Dim Arr() As Variant
    Dim R As Integer

    R = UBound(vector)
    ReDim Arr(R, width - 1)

    Dim i, j As Integer
    For i = 0 To R
        For j = 0 To width - 1
            Arr(i, j) = vector(i) ^ (j)
        Next j
    Next i

    Vander = Arr
End Function

' Function: Least2
' Perform a weighted least squares fit of order lambda of x to y
' using w as the weights
Public Function Least2(x As Variant, y As Variant, lambda As Integer, w As Variant)
    Dim nw As Integer
    nw = UBound(w)
    Dim Bigw() As Variant
    ReDim Bigw(nw + 3, n + 3)
    Bigw = Diagonal(w)
    Dim A, Aprime, V, Bigy, p, temp As Variant
    ReDim A(nw, lambda + 1)
    ReDim Aprime(lambda + 1, nw)
    ReDim V(lambda + 1, lambda + 1)
    ReDim Bigy(lambda + 1, 1)
    ReDim p(lambda + 1, 1)
    ReDim temp(lambda + 1, nw)
    A = Vander(x, lambda + 1)
    Aprime = Application.WorksheetFunction.transpose(A)
    temp = Application.WorksheetFunction.MMult(Aprime, Bigw)
    V = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MMult(Aprime, Bigw), A)
    y = Application.WorksheetFunction.transpose(y)
    Bigy = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MMult(Aprime, Bigw), y)
    p = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(V), Bigy)

    Least2 = p
End Function

' Function: Polyval
' Evaluate a polynomial at x given an array of the polynomial coefficients
' Parameters:
'    vector - An array of polynomial coefficients
'    x      - The value at which to evaluate the function
Private Function Polyval(vector As Variant, x As Variant)
    Dim length As Integer
    length = UBound(vector)
    Dim i As Integer
    Dim y As Variant
    y = 0
    For i = 1 To length
        y = y + vector(i) * x ^ (i - 1)
    Next i
    Polyval = y
End Function

' Function: Loess
' This function smooths a series of data using the loess algorithm.
' A polynomial of order lambda is fit to each data point and evaluated to smooth noisy data.
' alpha, the smoothness parameter is used to determine the width and weights at each point.
Public Function Loess(x As Variant, y As Variant, xnew As Variant, alpha As Variant, lambda As Integer)
    Dim n As Integer
    n = x.Rows.count
    n2 = xnew.Rows.count
    Dim q As Variant
    q = Application.WorksheetFunction.Floor(alpha * n, 1)
    q = Application.WorksheetFunction.Max(q, 1)
    q = Application.WorksheetFunction.Min(q, n)
    Dim z() As Variant
    ReDim z(n2)
    Dim i As Integer
    For i = 0 To n2 - 1
        Dim deltax() As Variant
        ReDim deltax(n)
        Dim j As Integer
 
        For j = 0 To n - 1
            deltax(j) = Abs(xnew(i + 1) - x(j + 1))
        Next j
        Dim qthdeltax As Variant
  
        qthdeltax = Application.WorksheetFunction.Small(deltax, q)
  
        Dim arg() As Variant
        ReDim arg(n)
        For j = 0 To n - 1
            arg(j) = Application.WorksheetFunction.Min(deltax(j) / (qthdeltax * Application.WorksheetFunction.Max(alpha, 1)), 1)
        Next j
        Dim tricube() As Variant
        ReDim tricube(n)
        Dim count As Integer
        count = 0
        For j = 0 To n - 1
            tricube(j) = (1 - Abs(arg(j)) ^ 3) ^ 3
            If tricube(j) > 0 Then count = count + 1
        Next j
        Dim weight() As Variant
        ReDim weight(count - 1)
        Dim count1, zeroes As Integer
        count1 = 0
        zeroes = 0
        For j = 0 To n - 1
            If tricube(j) > 0 Then
                weight(count1) = tricube(j)
                count1 = count1 + 1
            Else
                If count1 = 0 Then
                    zeroes = zeroes + 1
                End If
            End If
        Next j
        Dim weightx(), weighty() As Variant
        ReDim weightx(count - 1), weighty(count - 1)
        For j = 0 To count - 1
            weightx(j) = x(zeroes + j + 1)
            weighty(j) = y(zeroes + j + 1)
        Next j
        Dim p() As Variant
        ReDim p(lambda + 1, lambda + 1)
        p = Least2(weightx, weighty, lambda, weight)
        z(i) = Polyval(Application.WorksheetFunction.transpose(p), xnew(i + 1))
    Next i
    Loess = Application.WorksheetFunction.transpose(z)
End Function
