VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPolynomial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("Polynomial Evaluator")

Private m_coefficients() As Double

Public Function Initialize(coefficients() As Double) As CPolynomial

    m_coefficients = coefficients
    Set Initialize = Me
End Function

Public Property Get Degree() As Long
    Degree = UBound(m_coefficients) - LBound(m_coefficients)
End Property

Public Property Get Extent() As CArrayExtent
    Set Extents = CoreLibrary.Array1DExtent(m_coefficients)
End Property

Public Property Get A(ByVal i As Long) As Double
    A = m_coefficients(i)
End Property

Public Property Let A(ByVal i As Long, ByVal arg As Double)
    m_coefficients(i) = arg
End Property

Public Function EvaluateAt(ByVal x As Double) As Double
    EvaluateAt = ComputePolynomial(m_coefficients, x)
End Function

Public Sub EvaluateRange(x() As Double, results() As Double)

    Dim i As Long
    ReDim results(LBound(x) To UBound(x), 0 To 1) As Double
    For i = LBound(x) To UBound(x)
        Dim x_i As Double: x_i = x(i)
        results(i, 0) = x_i
        results(i, 1) = Me.EvaluateAt(x_i)
    Next i
End Sub

Public Function Evaluate(x() As Double) As Variant

    Dim yValues() As Double
    ReDim yValues(LBound(x, 1) To UBound(x, 1), 1 To 1)
    Dim i As Long
    For i = LBound(x, 1) To UBound(x, 1)
        yValues(i, 1) = Me.EvaluateAt(x(i, 1))
    Next i
    Evaluate = yValues

End Function

Public Sub EvaluateSequence(ByVal xStart As Double, ByVal xEnd As Double, ByVal nPoints As Long, results() As Double)

    Dim xStep As Double: xStep = (xEnd - xStart) / nPoints
    Dim i As Long
    ReDim results(0 To nPoints) As Double
    For i = 0 To nPoints
        Dim x_i As Double: x_i = xStart + i * xStep
        results(i, 0) = x_i
        results(i, 1) = Me.EvaluateAt(x_i)
    Next i
End Sub


Public Function FirstDerivative() As CPolynomial

    Dim dpdx() As Double
    If Me.Degree = 0 Then
        Set FirstDerivative = Nothing
        Exit Function
    End If
    
    ReDim dpdx(0 To Me.Degree - 1) As Double
    
    
    Dim i As Long
    For i = LBound(dpdx) To UBound(dpdx)
        dpdx(i) = (i + 1) * m_coefficients(i + 1)
    Next i
    
    Set FirstDerivative = ConstructPolynomial(dpdx)
    
End Function



Private Sub Class_Terminate()

    Erase m_coefficients
    
End Sub
