Attribute VB_Name = "TestPolynomials"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Polynomials")
Private Sub Test_EvaluatePolynomial_Degree1()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedValue As Double = 3#
    Dim coefficients() As Double
    ReDim coefficients(0 To 1) As Double
    coefficients(0) = 1#
    coefficients(1) = 2#
    Dim x As Double: x = 1#
    Dim computedValue As Double
    
    'Act:
    computedValue = ComputePolynomial(coefficients, x)
    
    
    'Assert:
    Assert.AreEqual expectedValue, computedValue, "P(1) = 2*1 + 1 = 3"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Polynomials")
Private Sub Test_EvaluatePolynomial_Degree2()
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedValue As Double = 17#
    Dim coefficients() As Double
    ReDim coefficients(0 To 2) As Double
    coefficients(0) = 1#
    coefficients(1) = 2#
    coefficients(2) = 3#
    Dim x As Double: x = 2#
    Dim computedValue As Double
    
    'Act:
    computedValue = ComputePolynomial(coefficients, x)
    
    
    'Assert:
    Assert.AreEqual expectedValue, computedValue, "P(1) =  3*2^2 + 2*2 + 1 = 3"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Polynomials")
Private Sub Test_EvaluatePolynomial_Degree2_List()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coefficients() As Double
    ReDim coefficients(0 To 2) As Double
    coefficients(0) = 1#
    coefficients(1) = 2#
    coefficients(2) = 3#
    Dim x() As Double
    ReDim x(0 To 9) As Double
    Dim i As Long
    
    For i = LBound(x) To UBound(x)
        x(i) = 1# * i
    Next i
    
    'Act:
    Debug.Print "Evaluate P(x) = " & coefficients(2) & " * x^2 + " & coefficients(1) & " * x + " & coefficients(0)
    For i = LBound(x) To UBound(x)
    
        Dim value As Double: value = ComputePolynomial(coefficients, x(i))
        Debug.Print "x = " & x(i) & ": P(x) = " & value
    Next i
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CPolynomial")
Private Sub Test_CPolynomial_Degree2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coefficients() As Double
    ReDim coefficients(0 To 2) As Double
    coefficients(0) = 1#
    coefficients(1) = 2#
    coefficients(2) = 3#
    Dim cPoly2 As CPolynomial: Set cPoly2 = ConstructPolynomial(coefficients)
    Dim x(10) As Double
    Dim i As Long
    For i = LBound(x) To UBound(x)
        x(i) = 1# * i
    Next i
    
    Dim results() As Double
    
    'Act:
    cPoly2.EvaluateRange x, results
    
    For i = LBound(x) To UBound(x)
        Debug.Print "p(" & results(i, 0) & ") = " & results(i, 1)
    Next i
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CPolynomial")
Private Sub Test_CPolynomial_Degree2_Derivative()
    On Error GoTo TestFail
    
    'Arrange:
    Dim coefficients() As Double
    ReDim coefficients(0 To 2) As Double
    coefficients(0) = 1#
    coefficients(1) = 2#
    coefficients(2) = 3#
    Dim cPoly2 As CPolynomial: Set cPoly2 = ConstructPolynomial(coefficients)
    Dim x(10) As Double
    Dim i As Long
    For i = LBound(x) To UBound(x)
        x(i) = 1# * i
    Next i
    
    Dim results() As Double
    Dim dPoly2 As CPolynomial
    
    'Act:
    Set dPoly2 = cPoly2.FirstDerivative()
    
    
    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

