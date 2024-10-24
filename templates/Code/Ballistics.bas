Attribute VB_Name = "Ballistics"
Option Explicit

Public Function TrajectoryInCM(ByVal elevationScopeUnits As Double, ByVal rangeInMeters As Double, ByVal scopeUnitsToMilsScale As Double) As Double

    Dim elevationMils As Double: elevationMils = elevationScopeUnits * scopeUnitsToMilsScale
    Dim elevationRadians As Double: elevationRadians = elevationMils / 1000#
    Dim elevationSign As Double: elevationSign = VBA.Math.Sgn(elevationRadians) * 1#
    Dim dropInM As Double: dropInM = rangeInMeters * VBA.Math.Tan(VBA.Math.Abs(elevationRadians)) * elevationSign
    
    TrajectoryInCM = dropInM * 100#   ' m * cm/m

End Function

Public Function TrajectoryToComeUp(data As Range, ByVal rangeInMeters As Double) As Double

    Dim dropCM As Double: dropCM = EvaluatePolynomial(data, rangeInMeters)
    Dim dropM As Double: dropM = dropCM / 100#
    
    Dim elevationMils As Double: elevationMils = 0#
    If Not rangeInMeters <= 0# Then
        elevationMils = 1000# * dropM / rangeInMeters
    End If
    
    TrajectoryToComeUp = -elevationMils

End Function


Public Function MOAToMRAD(ByVal moa As Double) As Double

    MOAToMRAD = (moa / 60#) * (Application.WorksheetFunction.Pi() / 180#) * 1000#

End Function

Public Function TrajectoryErrorCM(ByVal resolutionMILS As Double, ByVal scaleFactor As Double, ByVal rangeInM As Double) As Double

    Dim errorCM As Double: errorCM = (resolutionMILS / 1000#) * rangeInM * 100#
    TrajectoryErrorCM = errorCM * scaleFactor

End Function

'@Description("Calculate angular error, assuming a normal distribution with mu = 0, and specified precision")
' circularError1 is the high probability (one-sigma) inner circle
' circularError2 is the lower probability(two-sigma) outer circle
' sigma specifies the width of the Gaussian Distribution function
'   then tau = (1/sigma^2)
' and
'   F(x) = sqrt(tau/2Pi)*exp(-tau*x^2/2)
'
' the error is given by circularError1 * CDF(circularError1) + circularError2 * ( 1 - CDF(circularError1)
' that is, the inner circle has the highest probability of impact,
Public Function ProbableAngularPOIError(ByVal circularError1 As Double, ByVal circularError2, ByVal resolution As Double, Optional ByVal scaleFactor As Double = 1#) As Double

    Dim sigma As Double: sigma = resolution * scaleFactor
    Dim error1LB As Double: error1LB = -1# * circularError1 / 2#
    Dim error1UB As Double: error1UB = 1# * circularError1 / 2#
    
    Dim cdf1LB As Double: cdf1LB = Application.WorksheetFunction.NormDist(error1LB, 0#, sigma, True)
    Dim cdf1UB As Double: cdf1UB = Application.WorksheetFunction.NormDist(error1UB, 0, sigma, True)
    Dim weight1 As Double: weight1 = cdf1UB - cdf1LB
    
    Dim error2LB As Double: error2LB = -1# * circularError2 / 2#
    Dim error2UB As Double: error2UB = 1# * circularError2 / 2#
    Dim cdf2LB As Double: cdf2LB = Application.WorksheetFunction.NormDist(error2LB, 0#, sigma, True)
    Dim weight2 As Double: weight2 = cdf1LB - cdf2LB
    
    
    ProbableAngularPOIError = (weight1 * circularError1 + weight2 * circularError2)

End Function



Public Function MuzzleEnergyToMuzzleVelocityMPS(ByVal muzzleEnergyJoules As Double, ByVal bulletMassGrams) As Double

    Dim massKg As Double: massKg = bulletMassGrams / 1000#
    Dim muzzleVelocityMPS As Double: muzzleVelocityMPS = VBA.Math.Sqr(2# * muzzleEnergyJoules / massKg)
    MuzzleEnergyToMuzzleVelocityMPS = muzzleVelocityMPS
End Function

Public Function MuzzleVelocityToMuzzleEnergyJoules(ByVal velocityMPS As Double, ByVal bulletMassGrams) As Double

    Dim massKg As Double: massKg = bulletMassGrams / 1000#
    Dim muzzleEnergyJoules As Double: muzzleEnergyJoules = massKg * velocityMPS * velocityMPS / 2#
    MuzzleVelocityToMuzzleEnergyJoules = muzzleEnergyJoules
    
End Function

Public Function MuzzleEnergyToMuzzleVelocityFPS(ByVal muzzleEnergyFtLb As Double, ByVal bulletWeightGrains) As Double

    Dim bulletMassGrams As Double: bulletMassGrams = Application.WorksheetFunction.Convert(bulletWeightGrains, "grain", "g")
    Dim muzzleEnergyJoules As Double: muzzleEnergyJoules = Application.WorksheetFunction.Convert(muzzleEnergyFtLb, "flb", "J")
    Dim muzzleVelocityMPS As Double: muzzleVelocityMPS = MuzzleEnergyToMuzzleVelocityMPS(muzzleEnergyJoules, bulletMassGrams)
    Dim muzzleVelocityFPS As Double: muzzleVelocityFPS = Application.WorksheetFunction.Convert(muzzleVelocityMPS, "m", "ft") 'doesn't support direct m/sec -> ft/sec
    MuzzleEnergyToMuzzleVelocityFPS = muzzleVelocityFPS
End Function

Public Function MuzzleVelocityToMuzzleEnergyFtLb(ByVal muzzleVelocityFPS As Double, ByVal bulletWeightGrains As Double) As Double

    Dim bulletMassGrams As Double: bulletMassGrams = Application.WorksheetFunction.Convert(bulletWeightGrains, "grain", "g")
    Dim muzzleVelocityMPS As Double: muzzleVelocityMPS = Application.WorksheetFunction.Convert(muzzleVelocityFPS, "ft", "m")
    Dim muzzleEnergyJoules As Double: muzzleEnergyJoules = MuzzleVelocityToMuzzleEnergyJoules(muzzleVelocityMPS, bulletMassGrams)
    Dim muzzleEnergyFtLb As Double: muzzleEnergyFtLb = Application.WorksheetFunction.Convert(muzzleEnergyJoules, "J", "flb")
    MuzzleVelocityToMuzzleEnergyFtLb = muzzleEnergyFtLb

End Function



Public Function MuzzleVelocityToMuzzleEnergy(ByVal bulletWeight As Double, ByVal unitsWeight As String, ByVal muzzleVelocity As Double, ByVal velocityUnits As String, ByVal toUnits As String) As Double
    
    Dim bulletMassGrams As Double: bulletMassGrams = ConvertMassWeight(bulletWeight, unitsWeight, massUnitGrams)
    Dim muzzleVelocityMPS As Double: muzzleVelocityMPS = ConvertSpeed(muzzleVelocity, velocityUnits, speedUnitMPS)
    Dim muzzleEnergyJoules As Double: muzzleEnergyJoules = MuzzleVelocityToMuzzleEnergyJoules(muzzleVelocityMPS, bulletMassGrams)
    
    MuzzleVelocityToMuzzleEnergy = Application.WorksheetFunction.Convert(muzzleEnergyJoules, energyUnitJoules, toUnits)
End Function

'
'  generate a 2D Variant array that contains the point blank range results
'
Public Function PointBlankRangeListFromBallisticPolynomials(trajectoryCoefficients As Range, energyCoefficients As Range, ByVal targetCircleDiameterCM As Double, ByVal minRange As Double, ByVal rangeStep As Double, ByVal nSteps As Long) As Variant

    Dim trajectory As CPolynomial: Set trajectory = ConstructPolynomialFromRange(trajectoryCoefficients)
    Dim energyCurve As CPolynomial: Set energyCurve = ConstructPolynomialFromRange(energyCoefficients)
    Dim yMax As Double: yMax = targetCircleDiameterCM / 2#
    Dim yMin As Double: yMin = -1# * yMax
    Dim tracker As CTrajectoryTracker: Set tracker = TrajectoryTracker(nSteps, yMin, yMax)
    
    Dim pbrData As New ArrayList
    
    ' track the trajectory to look for the target crossings
    Dim x As Double: x = minRange
    Dim y As Double: y = 0#
    Dim ke As Double: ke = 0#
    Dim index As Long
    For index = 1 To nSteps
        y = trajectory.EvaluateAt(x)
        ke = energyCurve.EvaluateAt(x)
        tracker.TrackTrajectoryAndEnergy x, y, ke
        x = x + rangeStep
    Next index
    
    'now we can ask the trajectory tracker to find the target crossings, one by one
    Dim crossingList As New ArrayList
    tracker.FindCrossings crossingList
    
    'and evaluate the point blank range intervals
    Dim pbrList As New ArrayList
    tracker.BuildPointBlankRangeList crossingList, pbrList
    
    PointBlankRangeListFromBallisticPolynomials = OutputTable(pbrList)
    
End Function

Public Function PointBlankRangeListFromBallisticsSolution(trajectoryX As Range, trajectoryY As Range, energyRange As Range, ByVal targetCircleDiameterCM As Double) As Variant

    Dim yMax As Double: yMax = targetCircleDiameterCM / 2#
    Dim yMin As Double: yMin = -1# * yMax
    
    Dim trajectoryData() As Double: ConvertRangesToArray2D trajectoryX.value, trajectoryY.value, trajectoryData
    Dim energyData() As Double: ConvertRangesToArray2D trajectoryX.value, energyRange.value, energyData
    Dim nSteps As Long: nSteps = UBound(trajectoryData, 1) - LBound(trajectoryData, 1) + 1
    Dim tracker As CTrajectoryTracker: Set tracker = TrajectoryTracker(nSteps, yMin, yMax)
    
    Dim i As Long
    For i = LBound(trajectoryData, 1) To UBound(trajectoryData, 1)
        tracker.TrackTrajectoryAndEnergy trajectoryData(i, 1), trajectoryData(i, 2), energyData(i, 2)
    Next i
    
    'now we can ask the trajectory tracker to find the target crossings, one by one
    Dim crossingList As New ArrayList
    tracker.FindCrossings crossingList
    
    'and evaluate the point blank range intervals
    Dim pbrList As New ArrayList
    tracker.BuildPointBlankRangeList crossingList, pbrList
    
    PointBlankRangeListFromBallisticsSolution = OutputTable(pbrList)
    
    
End Function

Private Function OutputTable(pbrList As ArrayList) As Variant

    Dim output() As Variant
    ReDim output(1 To pbrList.Count + 1, 1 To 5)
    Dim i As Long
    output(1, 1) = "Interval #": output(1, 2) = "Minimum Range(m)": output(1, 3) = "Minimum Range Energy(J)": output(1, 4) = "Maximum Range(m)": output(1, 5) = "Maximum Range Energy(J)"
    For i = 1 To pbrList.Count
        Dim pbr As CPBRData: Set pbr = pbrList(i - 1)
        Dim row As Long: row = i + 1
        output(row, 1) = CStr(i): output(row, 2) = pbr.MinimumRange.value: output(row, 3) = pbr.EnergyAtMinimumRange.value: output(row, 4) = pbr.MaximumRange.value: output(row, 5) = pbr.EnergyAtMaximumRange.value
    Next i
    
    OutputTable = output

End Function






