Attribute VB_Name = "Ballistics"
Option Explicit

Public Function TrajectoryDropInCM(ByVal elevationScopeUnits As Double, ByVal rangeInMeters As Double, ByVal scopeUnitsToMilsScale As Double) As Double

    Dim elevationMils As Double: elevationMils = elevationScopeUnits * scopeUnitsToMilsScale
    Dim elevationRadians As Double: elevationRadians = elevationMils / 1000#
    Dim elevationSign As Double: elevationSign = VBA.Math.Sgn(elevationRadians) * 1#
    Dim dropInM As Double: dropInM = rangeInMeters * VBA.Math.Tan(VBA.Math.Abs(elevationRadians)) * elevationSign
    
    TrajectoryDropInCM = dropInM * 100#   ' m * cm/m

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


