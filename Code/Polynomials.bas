Attribute VB_Name = "Polynomials"
'@Description("Polynomial Evaluator: Horners Method")
'P[n](x; [a0, a1, a2, ..., a_n]) = a0 + a1*x + a2*x^2 + a3*x^3 + ... + a_n*x^n
'polyCoefficients: Range - the rows/column range containing the computed coefficients

Public Function EvaluatePolynomial(polyCoefficients As Range, ByVal x As Double) As Double

    Dim vals() As Variant: vals = polyCoefficients.value
    Dim a() As Double
    ReDim a(LBound(vals, 1) To UBound(vals, 1))
    Dim i As Long
    For i = LBound(vals, 1) To UBound(vals, 1)
        a(i) = CDbl(vals(i, 1))
    Next i
    
    
    EvaluatePolynomial = ComputePolynomial(a, x)
    
End Function

Public Function ComputePolynomial(a() As Double, ByVal x As Double) As Double

    ' compute polynomial value
    Dim i As Long
    Dim nUpper As Long: nUpper = UBound(a)
    Dim nLower As Long: nLower = LBound(a)
    Dim value As Double: value = a(nUpper)
    For i = nUpper - 1 To nLower Step -1
        value = x * value + a(i)
    Next i
    
    ComputePolynomial = value

End Function

'Perform Least Squares Regression Analysis on the data specified by the input range cells
' place the computed coefficients into the spreadsheet in the range specified by the output coefficients
'
Public Function ComputeLeastSquaresRegression(ByVal degree As Long, ByVal inputRangeX As Range, ByVal inputRangeY) As Variant
    If degree >= 1 Then
    
        Dim inputXData As Variant: inputXData = inputRangeX.value
        Dim inputYData As Variant: inputYData = inputRangeY.value
        Dim outputs As Variant
        Dim inputX() As Double, inputY() As Double
        ConvertRangeToArray inputXData, inputX
        ConvertRangeToArray inputYData, inputY
        
        Dim coefficients As Variant
        Dim rData As Variant
        
        If degree = 1 Then
        
            rData = FitPolynomial1(inputX, inputY)
        Else
            Dim powers() As Long
            ReDim powers(1 To degree) As Long
            
            Dim i As Long
            For i = 1 To degree
                powers(i) = i
            Next i
            
            rData = FitPolynomial(inputX, inputY, powers)
            
        End If
        
        coefficients = PolynomialCoefficients(rData, degree)
        ComputeLeastSquaresRegression = coefficients
    Else
        ComputeLeastSquaresRegression = ""
    End If
        
End Function

Public Function FitPolynomial1(inputX() As Double, inputY() As Double) As Variant

    Dim output As Variant: output = Application.LinEst(inputY, inputX, True, True)
    FitPolynomial1 = output
    
End Function



Public Function FitPolynomial(inputX() As Double, inputY() As Double, powers As Variant) As Variant
    FitPolynomial = Application.LinEst(inputY, Application.Power(inputX, powers), True, True)
End Function


Public Function PolynomialCoefficients(rData As Variant, ByVal degree As Long) As Variant

    Dim a() As Double
    ReDim a(1 To degree + 2, 1 To 1) As Double
    Dim dataIndex As Double: dataIndex = UBound(rData, 2)
    Dim i As Long
    a(1, 1) = rData(3, 1)
    For i = LBound(a) + 1 To UBound(a)
        a(i, 1) = rData(1, dataIndex)
        dataIndex = dataIndex - 1
    Next i
    
    PolynomialCoefficients = a
    
End Function

Private Sub ConvertRangeToArray(data As Variant, dataArray2D() As Double)

    ReDim dataArray2D(LBound(data, 1) To UBound(data, 1), LBound(data, 2) To UBound(data, 2))
    Dim row As Long, col As Long
    
    For row = LBound(data, 1) To UBound(data, 1)
        For col = LBound(data, 2) To UBound(data, 2)
            dataArray2D(row, col) = CDbl(data(row, col))
        Next col
    Next row
    
End Sub



Public Function ResizeRange(inRange As Range, ByVal lbRows As Long, ByVal lbCols As Long, ByVal ubRows As Long, ByVal ubCols As Long) As Range

    Dim outRows As Long: outRows = ubRows - lbRows + 1
    Dim outCols As Long: outCols = ubCols - lbCols + 1
    Dim newRange As Range: Set newRange = inRange.Resize(outRows, outCols)
    
    Set ResizeRange = newRange

End Function


