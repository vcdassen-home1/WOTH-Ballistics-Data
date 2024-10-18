Attribute VB_Name = "Conversions"
Option Explicit

Private Const PbDensity As Double = 11340#    ' kg/m^3
Private Const unitsInches As String = "in"
Private Const unitsMM As String = "mm"
Private Const massPb As Double = 0.45359237  ' one pound of lead in kilograms
Private Const mmToInch As Double = 0.0393700787
Private Const inchToMM As Double = 25.4
Private Const PI As Double = 3.1415926


Public Function ShotgunGaugeToBoreDiameter(ByVal gauge As Double, ByVal units As String) As Double

    Dim diameter As Double: diameter = 0#

    If units = unitsMM Then
        diameter = ShotgunGaugeToBoreDiameterMM(gauge)
    Else
        diameter = ShotgunGaugeToBoreDiameterInches(gauge)
    End If

    ShotgunGaugeToBoreDiameter = diameter

End Function

Public Function ShotgunGaugeToBoreDiameterMM(ByVal gauge As Double) As Double
    
    Dim diameterM3 As Double: diameterM3 = (6# * massPb) / (gauge * PbDensity * PI)
    Dim diameterMM As Double: diameterMM = Application.WorksheetFunction.Power(diameterM3, 1# / 3#) * 1000#
    ShotgunGaugeToBoreDiameterMM = diameterMM
    
End Function

Public Function ShotgunGaugeToBoreDiameterInches(ByVal gauge As Double) As Double

    ShotgunGaugeToBoreDiameterInches = ShotgunGaugeToBoreDiameterMM(gauge) * mmToInch

End Function
