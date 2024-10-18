Attribute VB_Name = "Polynomials"
'@Description("Polynomial Evaluator: Horners Method")
'P[n](x; [a0, a1, a2, ..., a_n]) = a0 + a1*x + a2*x^2 + a3*x^3 + ... + a_n*x^n
'polyCoefficients: Range - the rows/column range containing the computed coefficients

Public Function EvaluatePolynomial(polyCoefficients As Range, ByVal x As Double) As Double

    Dim A() As Double
    ConvertRangeToDoubleArray1D polyCoefficients, A
    EvaluatePolynomial = ComputePolynomial(A, x)
    
End Function

Public Sub ConvertRangeToDoubleArray1D(data As Range, array1D() As Double)

    Dim vals() As Variant: vals = data.value
    ReDim array1D(0 To UBound(vals, 1) - LBound(vals, 1)) As Double
    
    For i = LBound(vals, 1) To UBound(vals, 1)
        array1D(i - 1) = CDbl(vals(i, 1))
    Next i
    
End Sub

Public Function ComputePolynomial(A() As Double, ByVal x As Double) As Double

    ' compute polynomial value
    Dim i As Long
    Dim nUpper As Long: nUpper = UBound(A)
    Dim nLower As Long: nLower = LBound(A)
    
    Dim value As Double: value = A(nUpper)
    For i = nUpper - 1 To nLower Step -1
        value = x * value + A(i)
    Next i
    
    ComputePolynomial = value

End Function

'Perform Least Squares Regression Analysis on the data specified by the input range cells
' place the computed coefficients into the spreadsheet in the range specified by the output coefficients
'
Public Function ComputeLeastSquaresRegression(ByVal Degree As Long, ByVal inputRangeX As Range, ByVal inputRangeY) As Variant
    If Degree >= 1 Then
    
        Dim inputXData As Variant: inputXData = inputRangeX.value
        Dim inputYData As Variant: inputYData = inputRangeY.value
        Dim outputs As Variant
        Dim inputX() As Double, inputY() As Double
        ConvertRangeToArray2D inputXData, inputX
        ConvertRangeToArray2D inputYData, inputY
        
        Dim coefficients As Variant
        Dim rData As Variant
        
        If Degree = 1 Then
        
            rData = FitPolynomial1(inputX, inputY)
        Else
            Dim powers() As Long
            ReDim powers(1 To Degree) As Long
            
            Dim i As Long
            For i = 1 To Degree
                powers(i) = i
            Next i
            
            rData = FitPolynomial(inputX, inputY, powers)
            
        End If
        
        coefficients = PolynomialCoefficients(rData, Degree)
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


Public Function PolynomialCoefficients(rData As Variant, ByVal Degree As Long) As Variant

    Dim A() As Double
    ReDim A(1 To Degree + 2, 1 To 1) As Double
    Dim dataIndex As Double: dataIndex = UBound(rData, 2)
    Dim i As Long
    A(1, 1) = rData(3, 1)
    For i = LBound(A) + 1 To UBound(A)
        A(i, 1) = rData(1, dataIndex)
        dataIndex = dataIndex - 1
    Next i
    
    PolynomialCoefficients = A
    
End Function

Public Sub ConvertRangeToArray2D(data As Variant, dataArray2D() As Double)

    ' first find only the non-empty rows
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

Public Function ConstructPolynomialFromRange(polyCoefficients As Range) As CPolynomial

    ' skip empty rows
    Dim vals() As Variant: vals = polyCoefficients.value
    Dim A() As Double
    ReDim A(LBound(vals, 1) To UBound(vals, 1))
    Dim i As Long
    For i = LBound(vals, 1) To UBound(vals, 1)
        A(i) = CDbl(vals(i, 1))
    Next i
    
    Set ConstructPolynomialFromRange = ConstructPolynomial(A)
End Function

Public Function ConstructPolynomial(A() As Double) As CPolynomial

    Dim poly As New CPolynomial
    Set ConstructPolynomial = poly.Initialize(A)
End Function


Public Function FindPolynomialDegree2Roots(poly2 As CPolynomial, roots() As Double) As Boolean

    Dim a0 As Double: a0 = poly2.A(0)
    Dim a1 As Double: a1 = poly2.A(1)
    Dim a2 As Double: a2 = poly2.A(2)
    ReDim roots(0 To 1) As Double
    
    Dim discriminant As Double: discriminant = a1 * a1 - 4# * a2 * a0
    Dim r12 As Double: r12 = -a1 / (2 * a2)
    If discriminant >= 0 Then
        Dim r22 As Double: r22 = VBA.Math.Sqr(discriminant) / (2 * a2)
        roots(0) = r12 + r22
        roots(1) = r12 - r22
        FindPolynomialDegree2Roots = True
        Exit Function
    End If
    
    FindPolynomialDegree2Roots = False
    
End Function

Public Function FindPolynomialRoots(poly As CPolynomial, ByVal xStart As Double, ByVal xEnd As Double, ByVal xStep As Double, ByVal tolerance As Double, roots() As Double) As Boolean

    Dim dPoly As CPolynomial: Set dPoly = poly.FirstDerivative()
    If dPoly Is Nothing Then
        FindPolynomialRoots = False
        Exit Function
    End If
    
    ' bracket the zero crossings of the polynomial within the specified range
    ' use the midpoint of the bracketed range as the starting point for Newton-Raphson root finding
    ' we won't worry about complex roots here
    

End Function

Public Function GenerateRangeAddress(ByVal rStart As Double, ByVal rStep As Double, ByVal nSteps As Long, output As Range) As String
 
    ' calculate the output range
    Dim addrComponents() As String: addrComponents = Split(output.Address, "$")
    Dim outputCol As String: outputCol = addrComponents(1)
    Dim outputRow As Long: outputRow = CLng(addrComponents(2))
    Dim lastRow As Long: lastRow = outputRow + nSteps
    
    Dim builder As CStringBuilder: Set builder = StringBuilder(sDelimiter:="")
    
    With builder
        .Append output.Address, ":"
        .Append "$", outputCol
        .Append "$", CStr(lastRow)
    End With
     
    GenerateRangeAddress = builder.ToString()

End Function

Public Function GenerateRange(ByVal rStart As Double, ByVal rStep As Double, ByVal nSteps As Long) As Variant

    Dim xValues() As Double
    ReDim xValues(1 To nSteps + 1, 1 To 1) As Double
    Dim i As Long
    Dim x As Double
    For i = 1 To nSteps + 1
        x = rStart + rStep * (i - 1)
        xValues(i, 1) = x
    Next i
    
    GenerateRange = xValues


End Function

Public Function EvaluatePolynomialOverRange(polyCoefficients As Range, ByVal xValueRangeAddress As String) As Variant

    Dim poly As CPolynomial: Set poly = ConstructPolynomialFromRange(polyCoefficients)
    Dim xValues() As Double
    Dim xValueRange As Range: Set xValueRange = ActiveSheet.Range(xValueRangeAddress)
    ConvertRangeToArray2D xValueRange.value, xValues
    Dim yValues() As Double: yValues = poly.Evaluate(xValues)
    
    EvaluatePolynomialOverRange = yValues

End Function




