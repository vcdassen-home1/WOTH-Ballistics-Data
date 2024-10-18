Attribute VB_Name = "Test_LeastSquaresAlgorithms"
Option Explicit

Private Sub Test_LeastSquareFit_Degree1()

    Dim inputX() As Double, inputY() As Double
    ReDim inputX(1 To 10, 0) As Double
    ReDim inputY(1 To 10, 0) As Double
    
    Dim i As Long
    For i = LBound(inputX, 1) To UBound(inputX, 1)
        inputX(i, 0) = 1# * i
        inputY(i, 0) = 2# * i + 1#
    Next i

    Dim output As Variant
    output = FitPolynomial1(inputX, inputY)
    
    Debug.Print "done"

End Sub

Private Sub Test_LeastSquareFit_Degree2()

    Dim inputX() As Double, inputY() As Double
    ReDim inputX(1 To 10, 0) As Double
    ReDim inputY(1 To 10, 0) As Double
    
    Dim i As Long
    For i = LBound(inputX, 1) To UBound(inputX, 1)
        inputX(i, 0) = 1# * i
        inputY(i, 0) = 3# * i * i + 2# * i + 1#
    Next i

    Dim output As Variant
    output = FitPolynomial(inputX, inputY, Array(1, 2))
    
    Dim coefficients As Variant
    coefficients = PolynomialCoefficients(output, 2)
    
    Dim where As Range: Set where = Application.InputBox("Select Worksheet Range for output", Type:=8)
    
    'resize the range
    Dim here As Range: Set here = ResizeRange(where, LBound(coefficients, 1), 1, UBound(coefficients, 1), 1)
    here.value = coefficients

    
End Sub

Private Sub Test_WriteToSheet()



    
End Sub
